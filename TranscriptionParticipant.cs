namespace TeamsMediaBot;

/// <summary>Immutable identity attached to a media stream before any speech processing (Entra object id + Graph call participant id).</summary>
public sealed record TranscriptionParticipant(
    string ParticipantId,
    string DisplayName,
    string IntraId);
