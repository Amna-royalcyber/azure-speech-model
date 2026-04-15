using System.Collections.Concurrent;
using System.Text.Json;
using System.Threading;
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
using Microsoft.Extensions.Logging;

namespace TeamsMediaBot;

/// <summary>
/// One Azure Speech recognizer per media stream id. Identity is supplied by the caller (Graph + SSRC map); never inferred from audio.
/// </summary>
public sealed class AzureSpeechTranscriptionService : IAsyncDisposable
{
    private static readonly AudioStreamFormat Pcm16kMono = AudioStreamFormat.GetWaveFormatPCM(16000, 16, 1);

    private readonly BotSettings _settings;
    private readonly TranscriptBroadcaster _broadcaster;
    private readonly IChunkManager _chunkManager;
    private readonly ILogger<AzureSpeechTranscriptionService> _logger;
    private readonly ConcurrentDictionary<uint, StreamSession> _sessions = new();
    private volatile bool _disposed;
    private int _loggedMissingAzureConfig;

    private sealed class StreamSession
    {
        public readonly SemaphoreSlim Serialize = new(1, 1);
        public readonly object Gate = new();
        public PushAudioInputStream? Push;
        public SpeechRecognizer? Recognizer;
        public bool Started;
        public TranscriptionParticipant? Participant;
    }

    public AzureSpeechTranscriptionService(
        BotSettings settings,
        TranscriptBroadcaster broadcaster,
        IChunkManager chunkManager,
        ILogger<AzureSpeechTranscriptionService> logger)
    {
        _settings = settings;
        _broadcaster = broadcaster;
        _chunkManager = chunkManager;
        _logger = logger;
    }

    /// <summary>Process PCM for a stream with identity already resolved. Unknown SSRC must be dropped by the caller.</summary>
    public async Task ProcessAudioAsync(uint ssrc, TranscriptionParticipant participant, byte[] pcm16kMono, long timestampHns)
    {
        if (_disposed || pcm16kMono.Length == 0)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(_settings.AzureSpeechKey) ||
            string.IsNullOrWhiteSpace(_settings.AzureSpeechRegion))
        {
            if (Interlocked.Exchange(ref _loggedMissingAzureConfig, 1) == 0)
            {
                _logger.LogWarning(
                    "Azure Speech is not configured. Set Bot:AzureSpeechKey and Bot:AzureSpeechRegion (or env BOT_AZURE_SPEECH_KEY / BOT_AZURE_SPEECH_REGION).");
            }

            return;
        }

        var session = _sessions.GetOrAdd(ssrc, _ => new StreamSession());
        await session.Serialize.WaitAsync().ConfigureAwait(false);
        try
        {
            var shouldStart = false;
            lock (session.Gate)
            {
                if (!session.Started)
                {
                    session.Participant = participant;
                    shouldStart = true;
                }
            }

            if (shouldStart)
            {
                await StartRecognizerAsync(session, ssrc, participant).ConfigureAwait(false);
            }

            lock (session.Gate)
            {
                if (session.Push is not null)
                {
                    session.Push.Write(pcm16kMono);
                }
            }
        }
        finally
        {
            session.Serialize.Release();
        }
    }

    private async Task StartRecognizerAsync(StreamSession session, uint ssrc, TranscriptionParticipant participant)
    {
        lock (session.Gate)
        {
            if (session.Started)
            {
                return;
            }
        }

        try
        {
            var speechConfig = SpeechConfig.FromSubscription(_settings.AzureSpeechKey!, _settings.AzureSpeechRegion!);
            speechConfig.SpeechRecognitionLanguage = "en-US";

            var push = AudioInputStream.CreatePushStream(Pcm16kMono);
            var audioConfig = AudioConfig.FromStreamInput(push);
            var recognizer = new SpeechRecognizer(speechConfig, audioConfig);

            recognizer.Recognized += (_, e) =>
            {
                if (e.Result.Reason != ResultReason.RecognizedSpeech)
                {
                    return;
                }

                var text = e.Result.Text;
                if (string.IsNullOrWhiteSpace(text))
                {
                    return;
                }

                _logger.LogInformation("TRANSCRIPT [{DisplayName}]: {Text}", participant.DisplayName, text);
                var conf = TryParseConfidence(e.Result);
                _ = EmitTranscriptAsync(ssrc, participant, text, conf);
            };

            recognizer.Canceled += (_, e) =>
            {
                if (e.Reason == CancellationReason.Error)
                {
                    _logger.LogWarning("Azure Speech error on stream {SourceId}: {Details}", ssrc, e.ErrorDetails);
                }
            };

            await recognizer.StartContinuousRecognitionAsync().ConfigureAwait(false);

            lock (session.Gate)
            {
                session.Push = push;
                session.Recognizer = recognizer;
                session.Started = true;
                session.Participant = participant;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Azure Speech recognizer failed for stream {SourceId}.", ssrc);
        }
    }

    private async Task EmitTranscriptAsync(uint ssrc, TranscriptionParticipant participant, string text, double? confidence)
    {
        try
        {
            var utc = DateTime.UtcNow;
            await _broadcaster.BroadcastStructuredTranscriptAsync(
                participant.IntraId,
                participant.ParticipantId,
                participant.DisplayName,
                ssrc,
                text,
                confidence,
                utc).ConfigureAwait(false);

            var dedupeKey = $"{ssrc}|{utc.Ticks}|{text}";
            await _chunkManager.RecordFinalAsync(
                utc,
                participant.ParticipantId,
                participant.DisplayName,
                text,
                dedupeKey,
                ssrc).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Emit transcript failed for stream {SourceId}.", ssrc);
        }
    }

    public async ValueTask DisposeAsync()
    {
        _disposed = true;
        foreach (var kv in _sessions.ToArray())
        {
            await DisposeSessionAsync(kv.Value).ConfigureAwait(false);
        }

        _sessions.Clear();
    }

    private static async Task DisposeSessionAsync(StreamSession session)
    {
        SpeechRecognizer? rec;
        lock (session.Gate)
        {
            rec = session.Recognizer;
            session.Recognizer = null;
        }

        if (rec is not null)
        {
            try
            {
                await rec.StopContinuousRecognitionAsync().ConfigureAwait(false);
            }
            catch
            {
                // ignore
            }

            rec.Dispose();
        }

        lock (session.Gate)
        {
            session.Push?.Close();
            session.Push = null;
        }

        session.Serialize.Dispose();
    }

    private static double? TryParseConfidence(SpeechRecognitionResult result)
    {
        try
        {
            var json = result.Properties.GetProperty(PropertyId.SpeechServiceResponse_JsonResult);
            if (string.IsNullOrWhiteSpace(json))
            {
                return null;
            }

            using var doc = JsonDocument.Parse(json);
            if (!doc.RootElement.TryGetProperty("NBest", out var nBest) || nBest.GetArrayLength() == 0)
            {
                return null;
            }

            var first = nBest[0];
            return first.TryGetProperty("Confidence", out var c) ? c.GetDouble() : null;
        }
        catch
        {
            return null;
        }
    }
}
