using Microsoft.AspNetCore.SignalR;

namespace TeamsMediaBot;

public sealed class TranscriptHub : Hub
{
    // Marker hub for server-driven transcript and reconciliation events.
    // Clients should subscribe to: transcript, transcript-update, identity-resolved, transcript-retroactive-update, roster.
}
