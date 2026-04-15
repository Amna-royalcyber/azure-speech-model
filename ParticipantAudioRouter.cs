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
    private readonly ConcurrentDictionary<uint, DateTime> _unmappedSsrcFirstSeen = new();
    private static readonly TimeSpan BufferTimeout = TimeSpan.FromSeconds(15);
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
        _meetingParticipants.RegisterParticipantAudioRouter(this);
        _meetingParticipants.RegisterAudioRouterReconciler(FlushOrphanedAudio);
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
        MaybeRescanParticipantMediaStreams();

        if (!_ssrcMapper.HasMapping(ssrc) && !_participantManager.HasBinding(ssrc))
        {
            var firstSeen = _unmappedSsrcFirstSeen.GetOrAdd(ssrc, _ => DateTime.UtcNow);
            BufferOrphanedAudio(ssrc, rawPayload, timestampHns);
            if (!_unmappedSsrcLogThrottle.TryGetValue(ssrc, out var last) ||
                (DateTime.UtcNow - last) >= TimeSpan.FromSeconds(30))
            {
                _unmappedSsrcLogThrottle[ssrc] = DateTime.UtcNow;
                _logger.LogWarning(
                    "Buffering audio: SSRC/sourceId {Ssrc} is not mapped yet (buffer timeout {TimeoutSeconds}s).",
                    ssrc,
                    BufferTimeout.TotalSeconds);

                var orphanFor = DateTime.UtcNow - firstSeen;
                if (orphanFor >= TimeSpan.FromSeconds(10))
                {
                    _logger.LogError(
                        "ROUTER[ORPHAN] sourceId/SSRC {Ssrc} has been unmapped for {Seconds:F1}s. Mapping payload may be missing mediaStreams.",
                        ssrc,
                        orphanFor.TotalSeconds);
                }
            }
            return;
        }

        if (!_meetingParticipants.TryGetTranscriptionParticipant(ssrc, out var participant))
        {
            BufferOrphanedAudio(ssrc, rawPayload, timestampHns);
            _logger.LogWarning("ROUTER[MAP] SSRC/sourceId {Ssrc} mapped but participant details not resolved yet. Buffering until reconcile.", ssrc);
            return;
        }
        await FlushOrphanedAudio(ssrc, participant);

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

        await _azureSpeech.ProcessAudioAsync(ssrc, participant, pcm, timestampHns);
    }

    private void BufferOrphanedAudio(uint ssrc, byte[] rawPayload, long timestampHns)
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

        }
    }

    public async Task FlushOrphanedAudio(uint sourceId, TranscriptionParticipant identity)
    {
        if (!_audioBuffer.TryRemove(sourceId, out var queue))
        {
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
            sourceId,
            framesToReplay.Count,
            droppedExpired);

        foreach (var frame in framesToReplay)
        {
            await ProcessWithIdentity(sourceId, identity, frame.Payload, frame.TimestampHns);
        }
    }

    /// <summary>
    /// Late-binding entry point: once sourceId mapping is known, update speech identity and flush buffered frames.
    /// </summary>
    public async Task ReconcileSsrcAsync(uint sourceId)
    {
        TranscriptionParticipant? identity = null;
        if (_meetingParticipants.TryGetTranscriptionParticipant(sourceId, out var meetingIdentity))
        {
            identity = meetingIdentity;
        }
        else if (_participantManager.TryGetBinding(sourceId, out var binding) && binding is not null)
        {
            var fallbackId = string.IsNullOrWhiteSpace(binding.EntraOid)
                ? $"msi-pending-{sourceId}"
                : binding.EntraOid.Trim();
            var fallbackName = string.IsNullOrWhiteSpace(binding.DisplayName)
                ? _participantManager.GetTranscriptSpeakerLabel(sourceId)
                : binding.DisplayName.Trim();
            if (string.IsNullOrWhiteSpace(fallbackName))
            {
                fallbackName = fallbackId;
            }

            identity = new TranscriptionParticipant(fallbackId, fallbackName, fallbackId);
        }

        if (identity is null)
        {
            _logger.LogWarning("ROUTER[RECONCILE] sourceId/SSRC {Ssrc} mapped but identity not available yet.", sourceId);
            return;
        }

        await _azureSpeech.UpdateIdentityAsync(sourceId, identity).ConfigureAwait(false);
        await FlushBufferedAsync(sourceId, identity).ConfigureAwait(false);
    }

    private Task FlushBufferedAsync(uint sourceId, TranscriptionParticipant identity) =>
        FlushOrphanedAudio(sourceId, identity);

    private bool IsOrphan(uint ssrc)
    {
        if (!_unmappedSsrcFirstSeen.TryGetValue(ssrc, out var firstSeen))
        {
            return false;
        }

        return (DateTime.UtcNow - firstSeen) >= TimeSpan.FromSeconds(30);
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
            if ((DateTime.UtcNow - _lastParticipantRescanUtc) < TimeSpan.FromSeconds(2))
            {
                return;
            }

            _lastParticipantRescanUtc = DateTime.UtcNow;
        }

        TryHydrateFromCurrentRoster(call, botId);
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
        var sourceIds = GraphParticipantMediaStreams.ExtractSourceIds(resource, _logger).ToList();
        if (sourceIds.Count == 0 && !string.IsNullOrWhiteSpace(callPartId) && _attachedCall is not null)
        {
            // OnUpdated can provide partial delta resources; retry against the current participant object.
            try
            {
                foreach (var current in _attachedCall.Participants)
                {
                    if (!string.Equals(current.Resource?.Id, callPartId, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    sourceIds = GraphParticipantMediaStreams.ExtractSourceIds(current.Resource, _logger).ToList();
                    if (sourceIds.Count > 0)
                    {
                        _logger.LogInformation(
                            "MAP[RECONCILE] Recovered {Count} sourceIds for participant {ParticipantId} from live roster object.",
                            sourceIds.Count,
                            pid);
                    }
                    else
                    {
                        _logger.LogInformation(
                            "MAP[PAYLOAD] No mediaStreams yet for participant {ParticipantId}. Payload={Payload}",
                            pid,
                            GraphParticipantMediaStreams.BuildParticipantDiagnostics(current.Resource));
                    }
                    break;
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "MAP[RECONCILE] Could not inspect live roster object for participant {ParticipantId}.", pid);
            }
        }

        if (!sourceIds.Any())
        {
            _logger.LogWarning("No sourceIds found for participant {Name}", dn);
        }

        foreach (var sourceId in sourceIds)
        {
            if (!string.IsNullOrWhiteSpace(callPartId))
            {
                _meetingParticipants.BindMediaStreamToParticipant(sourceId, pid, callPartId);
            }

            _participantManager.TryBindAudioSource(sourceId, pid, dn, "Graph");
            _logger.LogInformation("[SSRC BIND] sourceId {SourceId} -> {DisplayName} ({ParticipantId})", sourceId, dn, pid);
            _unmappedSsrcLogThrottle.TryRemove(sourceId, out _);
            _unmappedSsrcFirstSeen.TryRemove(sourceId, out _);
            _ = _meetingParticipants.ReconcilePendingAudio(sourceId, pid);
        }
    }

    private void RemoveParticipantMappings(IParticipant participant)
    {
    }
}
