using Microsoft.AspNetCore.SignalR;
using Microsoft.Extensions.Logging;

namespace TeamsMediaBot;

public sealed class TranscriptBroadcaster
{
    private readonly IHubContext<TranscriptHub> _hubContext;
    private readonly ILogger<TranscriptBroadcaster> _logger;

    public TranscriptBroadcaster(
        IHubContext<TranscriptHub> hubContext,
        ILogger<TranscriptBroadcaster> logger)
    {
        _hubContext = hubContext;
        _logger = logger;
    }

    /// <summary>Forward final transcript as produced by the speech layer (identity already set upstream).</summary>
    public async Task BroadcastStructuredTranscriptAsync(
        string intraId,
        string participantId,
        string displayName,
        uint ssrc,
        string text,
        double? confidence,
        DateTime utteranceUtc)
    {
        try
        {
            await _hubContext.Clients.All.SendAsync(
                "transcript",
                new
                {
                    kind = "Final",
                    intraId,
                    participantId,
                    displayName,
                    speakerLabel = displayName,
                    sourceId = ssrc,
                    ssrc,
                    text,
                    confidence,
                    timestamp = new DateTimeOffset(utteranceUtc, TimeSpan.Zero),
                    azureAdObjectId = participantId,
                    tempLabel = false
                });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SignalR structured transcript broadcast failed for ssrc={Ssrc}.", ssrc);
        }
    }

    /// <summary>Optional UI hint when roster display name / Entra id updates for a stream. Does not change transcript identity (that is SSRC-bound before speech).</summary>
    public async Task BroadcastTranscriptIdentityUpdateAsync(uint sourceId, string? displayName, string? entraOid)
    {
        try
        {
            await _hubContext.Clients.All.SendAsync("transcript-update", new
            {
                type = "transcript-update",
                sourceId,
                displayName,
                azureAdObjectId = entraOid,
                timestamp = DateTimeOffset.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SignalR transcript-update failed for sourceId={SourceId}.", sourceId);
        }
    }

    public async Task BroadcastRosterAsync(IReadOnlyList<RosterParticipantDto> participants)
    {
        try
        {
            await _hubContext.Clients.All.SendAsync("roster", new
            {
                participants = participants.Select(p => new
                {
                    id = p.CallParticipantId,
                    displayName = p.DisplayName,
                    azureAdObjectId = p.AzureAdObjectId,
                    userPrincipalName = p.UserPrincipalName
                }).ToList(),
                timestamp = DateTimeOffset.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SignalR roster broadcast failed.");
        }
    }
}
