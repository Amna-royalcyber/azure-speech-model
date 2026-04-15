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
