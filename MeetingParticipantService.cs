using System.Collections.Concurrent;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Communications.Calls;
using Microsoft.Graph.Communications.Resources;
using Microsoft.Graph.Models;

namespace TeamsMediaBot;

/// <summary>
/// Tracks Teams meeting participants from Graph Communications roster updates and maps media stream ids
/// (<c>sourceId</c> from <c>mediaStreams</c>) to Entra identities before any transcription.
/// </summary>
public sealed class MeetingParticipantService
{
    private readonly TranscriptBroadcaster _broadcaster;
    private readonly IChunkManager _chunkManager;
    private readonly EntraUserResolver _entra;
    private readonly IParticipantManager _participantManager;
    private readonly SsrcParticipantMapper _ssrcMapper;
    private readonly ILogger<MeetingParticipantService> _logger;
    private readonly object _lock = new();
    private readonly ConcurrentDictionary<string, DateTime> _noSourceLogThrottle = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, DateTime> _payloadLogThrottle = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, DateTime> _consentBlockedLogThrottle = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Call participant resource ids (for removals).</summary>
    private readonly Dictionary<string, string> _callParticipantIdToAzureUserId = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Stable order of human participants for spk_N → row N mapping.</summary>
    private readonly List<RosterEntry> _rosterOrder = new();

    /// <summary>Teams audio <c>sourceId</c> → Entra object id from roster <c>mediaStreams</c> (links placeholders to real users).</summary>
    private readonly ConcurrentDictionary<uint, string> _audioSourceIdToAzureObjectId = new();

    /// <summary>Teams media <c>sourceId</c> → Graph call participant resource id (intra meeting id).</summary>
    private readonly ConcurrentDictionary<uint, string> _sourceIdToCallParticipantId = new();
    public event Action<uint, string>? SourceIdReconciled;
    private Func<uint, TranscriptionParticipant, Task>? _audioReconciler;
    private ParticipantAudioRouter? _participantAudioRouter;

    public MeetingParticipantService(
        TranscriptBroadcaster broadcaster,
        IChunkManager chunkManager,
        EntraUserResolver entra,
        IParticipantManager participantManager,
        SsrcParticipantMapper ssrcMapper,
        ILogger<MeetingParticipantService> logger)
    {
        _broadcaster = broadcaster;
        _chunkManager = chunkManager;
        _entra = entra;
        _participantManager = participantManager;
        _ssrcMapper = ssrcMapper;
        _logger = logger;
    }

    /// <summary>Register human participant as soon as Graph provides identity (does not wait for mediaStreams sourceId).</summary>
    public void AddOrUpdateParticipant(string participantId, string displayName, string intraId)
    {
        if (string.IsNullOrWhiteSpace(participantId))
        {
            return;
        }

        var pid = participantId.Trim();
        var dn = string.IsNullOrWhiteSpace(displayName) ? pid : displayName.Trim();
        var intra = string.IsNullOrWhiteSpace(intraId) ? pid : intraId.Trim();

        _participantManager.RegisterParticipant(pid, dn, DateTime.UtcNow);
        _logger.LogDebug("MAP[PARTICIPANT] Registered/updated participant {DisplayName} ({ParticipantId}), intra={IntraId}.", dn, pid, intra);

        lock (_lock)
        {
            var existingIdx = -1;
            for (var i = 0; i < _rosterOrder.Count; i++)
            {
                if (string.Equals(_rosterOrder[i].AzureAdObjectId, pid, StringComparison.OrdinalIgnoreCase))
                {
                    existingIdx = i;
                    break;
                }
            }

            if (existingIdx >= 0)
            {
                var cur = _rosterOrder[existingIdx];
                _rosterOrder[existingIdx] = cur with
                {
                    CallParticipantId = intra,
                    DisplayName = dn
                };
            }
            else
            {
                _rosterOrder.Add(new RosterEntry(intra, pid, dn, UserPrincipalName: null));
            }

            _callParticipantIdToAzureUserId[intra] = pid;
        }
    }

    public bool TryGetTranscriptionParticipant(uint sourceId, out TranscriptionParticipant participant)
    {
        participant = default!;
        if (!TryGetParticipantForMediaStream(sourceId, out var intraId, out var participantId, out var displayName))
        {
            return false;
        }

        participant = new TranscriptionParticipant(participantId, displayName, intraId);
        return true;
    }

    /// <summary>Graph + router: bind Teams media stream id to Entra user and call participant id (intra).</summary>
    public void BindMediaStreamToParticipant(uint sourceId, string entraObjectId, string intraCallParticipantId)
    {
        if (string.IsNullOrWhiteSpace(entraObjectId))
        {
            return;
        }

        var e = entraObjectId.Trim();
        var intra = string.IsNullOrWhiteSpace(intraCallParticipantId) ? e : intraCallParticipantId.Trim();
        _audioSourceIdToAzureObjectId[sourceId] = e;
        _sourceIdToCallParticipantId[sourceId] = intra;
        _ssrcMapper.Bind(sourceId, e);
        _logger.LogDebug("MAP[BIND] sourceId/SSRC {SourceId} -> participant {ParticipantId}, intra={IntraId}.", sourceId, e, intra);
        _ = _participantAudioRouter?.ReconcileSsrcAsync(sourceId);
    }

    /// <summary>
    /// Called when roster finally provides sourceId for an already ingested participant.
    /// Triggers downstream audio buffer flush + retroactive transcript identity update.
    /// </summary>
    public async Task ReconcilePendingAudio(uint sourceId, string participantId)
    {
        var intraId = _sourceIdToCallParticipantId.TryGetValue(sourceId, out var mappedIntraId) &&
                      !string.IsNullOrWhiteSpace(mappedIntraId)
            ? mappedIntraId
            : participantId;

        if (TryResolveAudioSourceToEntra(sourceId, out _, out var displayName) && !string.IsNullOrWhiteSpace(displayName))
        {
            var updated = await _chunkManager
                .ReconcileRecentIdentityAsync(sourceId, participantId, displayName, TimeSpan.FromSeconds(10))
                .ConfigureAwait(false);
            await _broadcaster.BroadcastTranscriptIdentityUpdateAsync(sourceId, displayName, participantId).ConfigureAwait(false);
            await _broadcaster.BroadcastIdentityResolved(sourceId, displayName, participantId).ConfigureAwait(false);
            if (updated > 0)
            {
                await _broadcaster.BroadcastTranscriptRetroactiveUpdateAsync(sourceId, displayName, participantId, updated).ConfigureAwait(false);
            }

            if (_audioReconciler is not null)
            {
                await _audioReconciler(
                    sourceId,
                    new TranscriptionParticipant(participantId, displayName, intraId)).ConfigureAwait(false);
            }
        }
        else
        {
            await _broadcaster.BroadcastTranscriptIdentityUpdateAsync(sourceId, participantId, participantId).ConfigureAwait(false);
            await _broadcaster.BroadcastIdentityResolved(sourceId, participantId, participantId).ConfigureAwait(false);
        }

        SourceIdReconciled?.Invoke(sourceId, participantId);
    }

    public void RegisterAudioRouterReconciler(Func<uint, TranscriptionParticipant, Task> reconciler)
    {
        _audioReconciler = reconciler;
    }

    public void RegisterParticipantAudioRouter(ParticipantAudioRouter participantAudioRouter)
    {
        _participantAudioRouter = participantAudioRouter;
    }

    public void AttachToCall(ICall call, string botAzureAdApplicationClientId)
    {
        _audioSourceIdToAzureObjectId.Clear();
        _sourceIdToCallParticipantId.Clear();
        _ssrcMapper.Clear();
        lock (_lock)
        {
            _callParticipantIdToAzureUserId.Clear();
            _rosterOrder.Clear();
        }

        var participants = call.Participants;
        participants.OnUpdated += (_, args) =>
        {
            try
            {
                var added = args.AddedResources ?? Array.Empty<IParticipant>();
                var updated = args.UpdatedResources ?? Array.Empty<IParticipant>();
                var removed = args.RemovedResources ?? Array.Empty<IParticipant>();

                _logger.LogInformation(
                    "GRAPH[PARTICIPANTS][DELTA] added={Added}, updated={Updated}, removed={Removed}.",
                    added.Count,
                    updated.Count,
                    removed.Count);

                foreach (var p in added)
                {
                    IngestParticipant(p, botAzureAdApplicationClientId);
                }

                foreach (var p in updated)
                {
                    IngestParticipant(p, botAzureAdApplicationClientId);
                }

                foreach (var p in removed)
                {
                    RemoveParticipant(p);
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Participant roster handler failed.");
            }
        };

        try
        {
            foreach (var p in participants)
            {
                IngestParticipant(p, botAzureAdApplicationClientId);
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Could not ingest existing call participants into roster.");
        }

        _logger.LogInformation("Subscribed to call participant roster updates; Entra profiles resolved via Microsoft Graph when needed.");
    }

    /// <summary>
    /// Deterministic identity for a media stream: requires Graph <c>mediaStreams[].sourceId</c> → user mapping.
    /// </summary>
    public bool TryGetParticipantForMediaStream(uint sourceId, out string intraId, out string participantId, out string displayName)
    {
        intraId = string.Empty;
        participantId = string.Empty;
        displayName = string.Empty;
        if (!_audioSourceIdToAzureObjectId.TryGetValue(sourceId, out var oid) || string.IsNullOrWhiteSpace(oid))
        {
            _logger.LogDebug("MAP[LOOKUP] sourceId/SSRC {SourceId} not mapped yet.", sourceId);
            return false;
        }

        participantId = oid.Trim();
        intraId = _sourceIdToCallParticipantId.TryGetValue(sourceId, out var callPid) ? callPid : string.Empty;
        if (TryResolveAudioSourceToEntra(sourceId, out _, out var dn) && !string.IsNullOrWhiteSpace(dn))
        {
            displayName = dn.Trim();
            return true;
        }

        displayName = participantId;
        return true;
    }

    /// <summary>Resolve Teams MSI/source id to Entra user using roster mediaStreams (when Graph lists source ids before audio bind upgrades).</summary>
    public bool TryResolveAudioSourceToEntra(uint sourceId, out string azureAdObjectId, out string displayName)
    {
        azureAdObjectId = string.Empty;
        displayName = string.Empty;
        if (!_audioSourceIdToAzureObjectId.TryGetValue(sourceId, out var oid) || string.IsNullOrWhiteSpace(oid))
        {
            return false;
        }

        azureAdObjectId = oid.Trim();
        lock (_lock)
        {
            foreach (var e in _rosterOrder)
            {
                if (string.Equals(e.AzureAdObjectId, azureAdObjectId, StringComparison.OrdinalIgnoreCase))
                {
                    displayName = string.IsNullOrWhiteSpace(e.DisplayName) ? azureAdObjectId : e.DisplayName.Trim();
                    return true;
                }
            }
        }

        displayName = azureAdObjectId;
        return true;
    }

    public IReadOnlyList<RosterParticipantDto> GetRosterSnapshot()
    {
        lock (_lock)
        {
            return _rosterOrder
                .Select(e => new RosterParticipantDto(
                    e.CallParticipantId,
                    e.DisplayName,
                    e.AzureAdObjectId,
                    e.UserPrincipalName))
                .ToList();
        }
    }

    private void IngestParticipant(IParticipant participant, string botClientId)
    {
        var resource = participant.Resource;
        if (resource is null)
        {
            return;
        }

        if (IsOurBot(resource, botClientId))
        {
            return;
        }

        var azureUserId = resource.Info?.Identity?.User?.Id;
        if (string.IsNullOrWhiteSpace(azureUserId))
        {
            return;
        }

        azureUserId = azureUserId.Trim();
        var callPartId = resource.Id;
        if (string.IsNullOrWhiteSpace(callPartId))
        {
            return;
        }

        var fromCall = resource.Info!.Identity!.User!.DisplayName?.Trim();
        var displayName = string.IsNullOrWhiteSpace(fromCall) ? null : fromCall;

        var needsGraph = string.IsNullOrWhiteSpace(displayName);

        _logger.LogDebug(
            "MAP[INGEST] Participant event: participantId={ParticipantId}, intra={IntraId}, displayName={DisplayName}, needsGraphEnrichment={NeedsGraph}.",
            azureUserId,
            callPartId,
            displayName ?? azureUserId,
            needsGraph);

        _logger.LogInformation(
            "GRAPH[PARTICIPANT][RESOURCE] id={ResourceId}, participant={ParticipantId}, displayName={DisplayName}, additionalDataKeys={Keys}.",
            callPartId,
            azureUserId,
            displayName ?? azureUserId,
            resource.AdditionalData is null ? "<none>" : string.Join(", ", resource.AdditionalData.Keys));

        AddOrUpdateParticipant(azureUserId, displayName ?? azureUserId, callPartId);

        var sourceIds = GraphParticipantMediaStreams.ExtractSourceIds(resource, _logger).ToList();
        _logger.LogDebug(
            "MAP[INGEST] Parsed {SourceCount} sourceIds for participant {ParticipantId}.",
            sourceIds.Count,
            azureUserId);
        foreach (var sid in sourceIds)
        {
            var isNewSourceForParticipant =
                !_audioSourceIdToAzureObjectId.TryGetValue(sid, out var existingParticipantId) ||
                !string.Equals(existingParticipantId, azureUserId, StringComparison.OrdinalIgnoreCase);

            BindMediaStreamToParticipant(sid, azureUserId, callPartId);
            _logger.LogInformation(
                "Authoritative stream map: sourceId {SourceId} -> {DisplayName} ({AzureAdObjectId}); intra={IntraId}.",
                sid,
                displayName ?? azureUserId,
                azureUserId,
                callPartId);

            if (isNewSourceForParticipant)
            {
                _logger.LogInformation(
                    "MAP[RECONCILE] New sourceId {SourceId} resolved for participant {ParticipantId}. Triggering pending audio reconciliation.",
                    sid,
                    azureUserId);
                _ = ReconcilePendingAudio(sid, azureUserId);
            }
        }

        if (sourceIds.Count == 0)
        {
            var mediaStreamsValueKind = ReadAdditionalDataValueKind(resource.AdditionalData, "mediaStreams");
            var mediaStreamsRaw = ReadAdditionalDataString(resource.AdditionalData, "mediaStreams");
            if (!string.IsNullOrWhiteSpace(mediaStreamsRaw) && mediaStreamsRaw.Length > 1200)
            {
                mediaStreamsRaw = mediaStreamsRaw[..1200] + "...<truncated>";
            }

            if (!_payloadLogThrottle.TryGetValue(azureUserId, out var payloadLogAt) ||
                (DateTime.UtcNow - payloadLogAt) >= TimeSpan.FromSeconds(30))
            {
                _payloadLogThrottle[azureUserId] = DateTime.UtcNow;
                _logger.LogInformation(
                    "MAP[PAYLOAD] Participant payload snapshot (safe): {Payload}",
                    GraphParticipantMediaStreams.BuildParticipantDiagnostics(resource));
            }

            var throttleKey = $"{azureUserId}:no-source";
            if (_noSourceLogThrottle.TryGetValue(throttleKey, out var last) && (DateTime.UtcNow - last) < TimeSpan.FromSeconds(30))
            {
                _ = PublishRosterAsync();
                if (needsGraph)
                {
                    _ = EnrichFromGraphAsync(azureUserId);
                }
                return;
            }
            _noSourceLogThrottle[throttleKey] = DateTime.UtcNow;

            var keys = resource.AdditionalData is null
                ? "<none>"
                : string.Join(", ", resource.AdditionalData.Keys);
            var anonymized = ReadAdditionalDataBool(resource.AdditionalData, "isIdentityAnonymized");
            var voiceConsent = ReadAdditionalDataString(resource.AdditionalData, "aiVoiceConsent");
            var aiVoiceConsentValue = ReadAiVoiceConsentValue(resource.AdditionalData);
            var rosterRole = ReadAdditionalDataString(resource.AdditionalData, "role");
            var endpointType = ReadAdditionalDataString(resource.AdditionalData, "endpointType");
            _logger.LogWarning(
                "GRAPH[STREAM][MISSING] Participant registered (Entra id {AzureAdObjectId}) but roster has no mediaStreams/sourceId yet; audio is buffered briefly until Graph publishes stream ids. additionalDataKeys={Keys}, mediaStreamsValueKind={MediaStreamsValueKind}, mediaStreamsRaw={MediaStreamsRaw}, isIdentityAnonymized={IsIdentityAnonymized}, aiVoiceConsent={AiVoiceConsent}, role={Role}, endpointType={EndpointType}.",
                azureUserId,
                keys,
                mediaStreamsValueKind ?? "<missing>",
                string.IsNullOrWhiteSpace(mediaStreamsRaw) ? "<missing>" : mediaStreamsRaw,
                anonymized.HasValue ? anonymized.Value.ToString() : "<missing>",
                string.IsNullOrWhiteSpace(voiceConsent) ? "<missing>" : voiceConsent,
                string.IsNullOrWhiteSpace(rosterRole) ? "<missing>" : rosterRole,
                string.IsNullOrWhiteSpace(endpointType) ? "<missing>" : endpointType);

            if (aiVoiceConsentValue == "0")
            {
                if (!_consentBlockedLogThrottle.TryGetValue(azureUserId, out var consentBlockedAt) ||
                    (DateTime.UtcNow - consentBlockedAt) >= TimeSpan.FromSeconds(30))
                {
                    _consentBlockedLogThrottle[azureUserId] = DateTime.UtcNow;
                    _logger.LogError(
                        "GRAPH[STREAM][BLOCKED_BY_CONSENT] aiVoiceConsentValue=0 for participant {AzureAdObjectId}. Graph will not publish mediaStreams/sourceId; SSRC cannot be identity-mapped. Fix tenant/meeting voice consent and rejoin call.",
                        azureUserId);
                }
            }
        }

        _ = PublishRosterAsync();

        if (needsGraph)
        {
            _ = EnrichFromGraphAsync(azureUserId);
        }
    }

    private async Task EnrichFromGraphAsync(string azureUserId)
    {
        try
        {
            _logger.LogDebug("MAP[GRAPH] Enriching participant profile for {ParticipantId}.", azureUserId);
            var profile = await _entra.GetUserAsync(azureUserId).ConfigureAwait(false);
            if (profile is null)
            {
                _logger.LogDebug("MAP[GRAPH] No profile found for {ParticipantId}.", azureUserId);
                return;
            }

            var dn = string.IsNullOrWhiteSpace(profile.DisplayName) ? profile.Id : profile.DisplayName.Trim();
            var upn = profile.UserPrincipalName;

            lock (_lock)
            {
                for (var i = 0; i < _rosterOrder.Count; i++)
                {
                    if (!string.Equals(_rosterOrder[i].AzureAdObjectId, azureUserId, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    _rosterOrder[i] = _rosterOrder[i] with
                    {
                        DisplayName = dn,
                        UserPrincipalName = string.IsNullOrWhiteSpace(upn) ? _rosterOrder[i].UserPrincipalName : upn.Trim()
                    };
                }
            }

            await PublishRosterAsync().ConfigureAwait(false);
            _logger.LogDebug("MAP[GRAPH] Enrichment applied for {ParticipantId}.", azureUserId);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Graph enrichment failed for {UserId}.", azureUserId);
        }
    }

    private void RemoveParticipant(IParticipant participant)
    {
        var callPartId = participant.Resource?.Id;
        if (string.IsNullOrWhiteSpace(callPartId))
        {
            return;
        }

        string? removedAzureId = null;
        lock (_lock)
        {
            if (!_callParticipantIdToAzureUserId.TryGetValue(callPartId, out removedAzureId))
            {
                return;
            }

            _callParticipantIdToAzureUserId.Remove(callPartId);
            _rosterOrder.RemoveAll(e => string.Equals(e.CallParticipantId, callPartId, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(removedAzureId))
        {
            foreach (var kv in _audioSourceIdToAzureObjectId.ToArray())
            {
                if (string.Equals(kv.Value, removedAzureId, StringComparison.OrdinalIgnoreCase))
                {
                    _audioSourceIdToAzureObjectId.TryRemove(kv.Key, out _);
                    _sourceIdToCallParticipantId.TryRemove(kv.Key, out _);
                    _ssrcMapper.RemoveSsrc(kv.Key);
                    _logger.LogInformation(
                        "MAP[REMOVE] Removed sourceId/SSRC {SourceId} mapping for participant {ParticipantId}.",
                        kv.Key,
                        removedAzureId);
                }
            }
        }

        _ = PublishRosterAsync();
    }

    private async Task PublishRosterAsync()
    {
        List<RosterParticipantDto> snapshot;
        lock (_lock)
        {
            snapshot = _rosterOrder
                .Select(e => new RosterParticipantDto(
                    e.CallParticipantId,
                    e.DisplayName,
                    e.AzureAdObjectId,
                    e.UserPrincipalName))
                .ToList();
        }

        await _broadcaster.BroadcastRosterAsync(snapshot).ConfigureAwait(false);
    }

    private static bool IsOurBot(Participant resource, string botClientId)
    {
        var appId = resource.Info?.Identity?.Application?.Id;
        return !string.IsNullOrEmpty(appId) &&
               string.Equals(appId.Trim(), botClientId.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static bool? ReadAdditionalDataBool(IDictionary<string, object>? additionalData, string key)
    {
        if (additionalData is null)
        {
            return null;
        }

        foreach (var kv in additionalData)
        {
            if (!string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var value = kv.Value;
            if (value is null)
            {
                return null;
            }

            if (value is bool b)
            {
                return b;
            }

            if (value is JsonElement je)
            {
                if (je.ValueKind is JsonValueKind.True or JsonValueKind.False)
                {
                    return je.GetBoolean();
                }

                if (je.ValueKind == JsonValueKind.String && bool.TryParse(je.GetString(), out var jb))
                {
                    return jb;
                }
            }

            if (bool.TryParse(Convert.ToString(value), out var parsed))
            {
                return parsed;
            }

            return null;
        }

        return null;
    }

    private static string? ReadAdditionalDataString(IDictionary<string, object>? additionalData, string key)
    {
        if (additionalData is null)
        {
            return null;
        }

        foreach (var kv in additionalData)
        {
            if (!string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var value = kv.Value;
            if (value is null)
            {
                return null;
            }

            if (value is JsonElement je)
            {
                if (je.ValueKind == JsonValueKind.String)
                {
                    return je.GetString();
                }

                return je.GetRawText();
            }

            return Convert.ToString(value);
        }

        return null;
    }

    private static string? ReadAdditionalDataValueKind(IDictionary<string, object>? additionalData, string key)
    {
        if (additionalData is null)
        {
            return null;
        }

        foreach (var kv in additionalData)
        {
            if (!string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (kv.Value is null)
            {
                return "null";
            }

            if (kv.Value is JsonElement je)
            {
                return je.ValueKind.ToString();
            }

            return kv.Value.GetType().Name;
        }

        return null;
    }

    private static string? ReadAiVoiceConsentValue(IDictionary<string, object>? additionalData)
    {
        var raw = ReadAdditionalDataString(additionalData, "aiVoiceConsent");
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        try
        {
            using var doc = JsonDocument.Parse(raw);
            var root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                return null;
            }

            if (!root.TryGetProperty("value", out var valueNode) || valueNode.ValueKind != JsonValueKind.Object)
            {
                return null;
            }

            if (!valueNode.TryGetProperty("aiVoiceConsentValue", out var consentNode))
            {
                return null;
            }

            if (consentNode.ValueKind == JsonValueKind.String)
            {
                return consentNode.GetString();
            }

            if (consentNode.ValueKind == JsonValueKind.Number)
            {
                return consentNode.GetRawText();
            }

            return null;
        }
        catch
        {
            return null;
        }
    }

    private readonly record struct RosterEntry(
        string CallParticipantId,
        string AzureAdObjectId,
        string DisplayName,
        string? UserPrincipalName);
}

public sealed record RosterParticipantDto(
    string CallParticipantId,
    string DisplayName,
    string AzureAdObjectId,
    string? UserPrincipalName);
