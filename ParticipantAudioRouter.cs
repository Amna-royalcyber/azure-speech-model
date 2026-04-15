using System.Collections.Concurrent;
using System.Threading;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Communications.Calls;
using Microsoft.Graph.Communications.Resources;
using Microsoft.Skype.Bots.Media;

namespace TeamsMediaBot;

/// <summary>
/// Routes each frame by SSRC (Teams <c>sourceId</c>) after Graph has bound that id to a participant. Unmapped SSRC → audio dropped.
/// </summary>
public sealed class ParticipantAudioRouter
{
    private sealed record BufferedFrame(byte[] Payload, long TimestampHns, DateTime EnqueuedUtc);

    private readonly AudioProcessor _audioProcessor;
    private readonly AzureSpeechTranscriptionService _azureSpeech;
    private readonly MeetingParticipantService _meetingParticipants;
    private readonly SsrcParticipantMapper _ssrcMapper;
    private readonly ParticipantManager _participantManager;
    private readonly ILogger<ParticipantAudioRouter> _logger;

    private ICall? _attachedCall;
    private string _botClientId = string.Empty;
    private readonly ConcurrentDictionary<uint, DateTime> _unmappedSsrcLogThrottle = new();
    private readonly ConcurrentDictionary<uint, Queue<BufferedFrame>> _audioBuffer = new();
    private static readonly TimeSpan BufferTimeout = TimeSpan.FromSeconds(12);
    private const int MaxBufferedFramesPerSsrc = 120;

    public ParticipantAudioRouter(
        AudioProcessor audioProcessor,
        AzureSpeechTranscriptionService azureSpeech,
        MeetingParticipantService meetingParticipants,
        SsrcParticipantMapper ssrcMapper,
        ParticipantManager participantManager,
        ILogger<ParticipantAudioRouter> logger)
    {
        _audioProcessor = audioProcessor;
        _azureSpeech = azureSpeech;
        _meetingParticipants = meetingParticipants;
        _ssrcMapper = ssrcMapper;
        _participantManager = participantManager;
        _logger = logger;
    }

    public void AttachToCall(ICall call, string botClientId)
    {
        _attachedCall = call;
        _botClientId = botClientId ?? string.Empty;

        var bot = _botClientId;
        call.Participants.OnUpdated += (_, args) =>
        {
            foreach (var p in args.AddedResources)
            {
                UpsertParticipantMappings(p, bot);
            }

            foreach (var p in args.UpdatedResources)
            {
                UpsertParticipantMappings(p, bot);
            }

            foreach (var p in args.RemovedResources)
            {
                RemoveParticipantMappings(p);
            }
        };

        TryHydrateFromCurrentRoster(call, bot);
    }

    private void TryHydrateFromCurrentRoster(ICall call, string botClientId)
    {
        try
        {
            foreach (var p in call.Participants)
            {
                UpsertParticipantMappings(p, botClientId);
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Could not hydrate participant source bindings from current roster.");
        }
    }

    /// <summary>
    /// <paramref name="ssrc"/> is the Teams media stream id for this frame (same as Graph <c>mediaStreams[].sourceId</c>).
    /// </summary>
    public async Task HandleAudioAsync(uint ssrc, byte[] rawPayload, long timestampHns)
    {
        _logger.LogDebug(
            "ROUTER[RECV] SSRC/sourceId {Ssrc}, bytes={Bytes}, ts={TimestampHns}.",
            ssrc,
            rawPayload.Length,
            timestampHns);

        if (!_ssrcMapper.HasMapping(ssrc))
        {
            BufferUnmappedAudio(ssrc, rawPayload, timestampHns);
            if (!_unmappedSsrcLogThrottle.TryGetValue(ssrc, out var last) ||
                (DateTime.UtcNow - last) >= TimeSpan.FromSeconds(30))
            {
                _unmappedSsrcLogThrottle[ssrc] = DateTime.UtcNow;
                _logger.LogWarning(
                    "Buffering audio: SSRC/sourceId {Ssrc} is not mapped yet (buffer timeout {TimeoutSeconds}s).",
                    ssrc,
                    BufferTimeout.TotalSeconds);
                Console.WriteLine($"[CONSOLE][ROUTER][BUFFERING] sourceId={ssrc}, timeoutSec={BufferTimeout.TotalSeconds}, bytes={rawPayload.Length}");
            }
            return;
        }

        _logger.LogDebug("ROUTER[MAP] SSRC/sourceId {Ssrc} has mapping.", ssrc);
        if (!_meetingParticipants.TryGetTranscriptionParticipant(ssrc, out var participant))
        {
            _logger.LogWarning("ROUTER[MAP] SSRC/sourceId {Ssrc} mapped but participant details not resolved yet.", ssrc);
            return;
        }

        if (_audioBuffer.TryRemove(ssrc, out var bufferedFrames))
        {
            var framesToReplay = new List<BufferedFrame>();
            lock (bufferedFrames)
            {
                while (bufferedFrames.Count > 0)
                {
                    var buffered = bufferedFrames.Dequeue();
                    if ((DateTime.UtcNow - buffered.EnqueuedUtc) > BufferTimeout)
                    {
                        continue;
                    }

                    framesToReplay.Add(buffered);
                }
            }

            _logger.LogInformation(
                "ROUTER[FLUSH] Replaying {Count} buffered frames for SSRC/sourceId {Ssrc}.",
                framesToReplay.Count,
                ssrc);
            Console.WriteLine($"[CONSOLE][ROUTER][FLUSH] sourceId={ssrc}, replayFrames={framesToReplay.Count}");
            foreach (var buffered in framesToReplay)
            {
                await ProcessWithIdentity(ssrc, participant, buffered.Payload, buffered.TimestampHns);
            }
        }

        await ProcessWithIdentity(ssrc, participant, rawPayload, timestampHns);
    }

    private async Task ProcessWithIdentity(uint ssrc, TranscriptionParticipant participant, byte[] rawPayload, long timestampHns)
    {
        var pcm = _audioProcessor.ConvertToPcm(new AudioFrame(
            Data: rawPayload,
            Timestamp: timestampHns,
            Length: rawPayload.Length,
            Format: AudioFormat.Pcm16K));

        if (pcm.Length == 0)
        {
            _logger.LogDebug("ROUTER[PCM] Converted PCM empty for SSRC/sourceId {Ssrc}; frame dropped.", ssrc);
            return;
        }

        _logger.LogDebug(
            "ROUTER[PCM] SSRC/sourceId {Ssrc} -> participant {DisplayName} ({ParticipantId}), pcmBytes={Bytes}.",
            ssrc,
            participant.DisplayName,
            participant.ParticipantId,
            pcm.Length);
        await _azureSpeech.ProcessAudioAsync(ssrc, participant, pcm, timestampHns);
    }

    private void BufferUnmappedAudio(uint ssrc, byte[] rawPayload, long timestampHns)
    {
        var queue = _audioBuffer.GetOrAdd(ssrc, _ => new Queue<BufferedFrame>());
        var droppedForCapacity = 0;
        var droppedForAge = 0;
        lock (queue)
        {
            queue.Enqueue(new BufferedFrame(rawPayload, timestampHns, DateTime.UtcNow));
            while (queue.Count > MaxBufferedFramesPerSsrc)
            {
                queue.Dequeue();
                droppedForCapacity++;
            }
            while (queue.Count > 0 && (DateTime.UtcNow - queue.Peek().EnqueuedUtc) > BufferTimeout)
            {
                queue.Dequeue();
                droppedForAge++;
            }

            _logger.LogDebug(
                "ROUTER[BUFFER] SSRC/sourceId {Ssrc}: queued frame, depth={Depth}, droppedCapacity={DroppedCapacity}, droppedAge={DroppedAge}.",
                ssrc,
                queue.Count,
                droppedForCapacity,
                droppedForAge);
        }
    }

    private async Task FlushBufferedAsync(uint ssrc)
    {
        if (!_audioBuffer.TryRemove(ssrc, out var queue))
        {
            _logger.LogDebug("ROUTER[FLUSH] No buffered audio for SSRC/sourceId {Ssrc}.", ssrc);
            return;
        }

        var framesToReplay = new List<BufferedFrame>();
        var droppedExpired = 0;
        lock (queue)
        {
            while (queue.Count > 0)
            {
                var frame = queue.Dequeue();
                if ((DateTime.UtcNow - frame.EnqueuedUtc) <= BufferTimeout)
                {
                    framesToReplay.Add(frame);
                }
                else
                {
                    droppedExpired++;
                }
            }
        }

        _logger.LogInformation(
            "ROUTER[FLUSH] SSRC/sourceId {Ssrc}: replaying={ReplayCount}, droppedExpired={DroppedExpired}.",
            ssrc,
            framesToReplay.Count,
            droppedExpired);

        foreach (var frame in framesToReplay)
        {
            if (!_meetingParticipants.TryGetTranscriptionParticipant(ssrc, out var participant))
            {
                _logger.LogWarning("ROUTER[FLUSH] SSRC/sourceId {Ssrc}: participant could not be resolved during replay.", ssrc);
                break;
            }

            await ProcessWithIdentity(ssrc, participant, frame.Payload, frame.TimestampHns);
        }
    }

    private void UpsertParticipantMappings(IParticipant participant, string botClientId)
    {
        var resource = participant.Resource;
        var identity = resource?.Info?.Identity;
        var appId = identity?.Application?.Id;
        if (!string.IsNullOrWhiteSpace(appId) &&
            string.Equals(appId.Trim(), botClientId, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var participantId = identity?.User?.Id;
        if (string.IsNullOrWhiteSpace(participantId))
        {
            return;
        }

        var displayName = identity?.User?.DisplayName;
        if (string.IsNullOrWhiteSpace(displayName))
        {
            displayName = participantId;
        }

        var pid = participantId.Trim();
        var dn = displayName.Trim();
        _participantManager.RegisterParticipant(pid, dn, DateTime.UtcNow);
        _logger.LogDebug("ROUTER[ROSTER] Participant update: {DisplayName} ({ParticipantId}).", dn, pid);

        var callPartId = resource?.Id;
        foreach (var sourceId in GraphParticipantMediaStreams.ExtractSourceIds(resource))
        {
            if (!string.IsNullOrWhiteSpace(callPartId))
            {
                _meetingParticipants.BindMediaStreamToParticipant(sourceId, pid, callPartId);
            }

            _participantManager.TryBindAudioSource(sourceId, pid, dn, "Graph");
            _logger.LogInformation("[SSRC BIND] sourceId {SourceId} -> {DisplayName} ({ParticipantId})", sourceId, dn, pid);
            Console.WriteLine($"[CONSOLE][ROUTER][BIND] sourceId={sourceId}, participant={pid}, displayName={dn}");
            _unmappedSsrcLogThrottle.TryRemove(sourceId, out _);
            _ = FlushBufferedAsync(sourceId);
        }
    }

    private void RemoveParticipantMappings(IParticipant participant)
    {
    }
}
