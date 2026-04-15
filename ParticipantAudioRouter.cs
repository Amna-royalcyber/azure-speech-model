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
    private readonly object _rescanLock = new();
    private DateTime _lastParticipantRescanUtc = DateTime.MinValue;
    private readonly ConcurrentDictionary<uint, DateTime> _unmappedSsrcLogThrottle = new();
    private readonly ConcurrentDictionary<uint, Queue<BufferedFrame>> _audioBuffer = new();
    private static readonly TimeSpan BufferTimeout = TimeSpan.FromSeconds(3);

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
        lock (_rescanLock)
        {
            _lastParticipantRescanUtc = DateTime.MinValue;
        }

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
        MaybeRescanParticipantMediaStreams();

        if (!_ssrcMapper.HasMapping(ssrc))
        {
            BufferUnmappedAudio(ssrc, rawPayload, timestampHns);
            if (!_unmappedSsrcLogThrottle.TryGetValue(ssrc, out var last) ||
                (DateTime.UtcNow - last) >= TimeSpan.FromSeconds(10))
            {
                _unmappedSsrcLogThrottle[ssrc] = DateTime.UtcNow;
                _logger.LogWarning(
                    "Buffering audio: SSRC/sourceId {Ssrc} is not mapped yet (buffer timeout {TimeoutSeconds}s).",
                    ssrc,
                    BufferTimeout.TotalSeconds);
            }
            return;
        }

        if (!_meetingParticipants.TryGetTranscriptionParticipant(ssrc, out var participant))
        {
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
            return;
        }

        _logger.LogDebug("PCM for SSRC {Ssrc} ({Name}).", ssrc, participant.DisplayName);
        await _azureSpeech.ProcessAudioAsync(ssrc, participant, pcm, timestampHns);
    }

    private void BufferUnmappedAudio(uint ssrc, byte[] rawPayload, long timestampHns)
    {
        var queue = _audioBuffer.GetOrAdd(ssrc, _ => new Queue<BufferedFrame>());
        lock (queue)
        {
            queue.Enqueue(new BufferedFrame(rawPayload, timestampHns, DateTime.UtcNow));
            while (queue.Count > 0 && (DateTime.UtcNow - queue.Peek().EnqueuedUtc) > BufferTimeout)
            {
                queue.Dequeue();
            }
        }
    }

    private void MaybeRescanParticipantMediaStreams()
    {
        var call = _attachedCall;
        var botId = _botClientId;
        if (call is null || string.IsNullOrWhiteSpace(botId))
        {
            return;
        }

        lock (_rescanLock)
        {
            if ((DateTime.UtcNow - _lastParticipantRescanUtc).TotalSeconds < 2.5)
            {
                return;
            }

            _lastParticipantRescanUtc = DateTime.UtcNow;
        }

        try
        {
            _meetingParticipants.ResyncParticipantMediaStreamsFromCall(call, botId);
            foreach (var p in call.Participants)
            {
                UpsertParticipantMappings(p, botId);
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Periodic participant mediaStreams rescan failed.");
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

        var callPartId = resource?.Id;
        foreach (var sourceId in GraphParticipantMediaStreams.ExtractSourceIds(resource))
        {
            if (!string.IsNullOrWhiteSpace(callPartId))
            {
                _meetingParticipants.BindMediaStreamToParticipant(sourceId, pid, callPartId);
            }

            _participantManager.TryBindAudioSource(sourceId, pid, dn, "Graph");
            _logger.LogInformation("[SSRC BIND] sourceId {SourceId} -> {DisplayName} ({ParticipantId})", sourceId, dn, pid);
        }
    }

    private void RemoveParticipantMappings(IParticipant participant)
    {
    }
}
