using System.Threading;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Communications.Calls;
using Microsoft.Graph.Communications.Calls.Media;
using Microsoft.Graph.Communications.Client;
using Microsoft.Skype.Bots.Media;

namespace TeamsMediaBot;

public sealed class MediaHandler
{
    private readonly ILogger<MediaHandler> _logger;
    private readonly ParticipantAudioRouter _participantAudioRouter;
    private IAudioSocket? _audioSocket;
    private int _loggedNoUnmixed;

    public MediaHandler(
        ILogger<MediaHandler> logger,
        ParticipantAudioRouter participantAudioRouter)
    {
        _logger = logger;
        _participantAudioRouter = participantAudioRouter;
    }

    public IMediaSession CreateMediaSession(ICommunicationsClient communicationsClient)
    {
        var mediaConfiguration = new AudioSocketSettings
        {
            StreamDirections = StreamDirection.Recvonly,
            SupportedAudioFormat = AudioFormat.Pcm16K,
            ReceiveUnmixedMeetingAudio = true,
            EnableAudioHealingForUnmixed = true
        };

        var mediaSession = communicationsClient.CreateMediaSession(
            audioSocketSettings: mediaConfiguration,
            videoSocketSettings: (IEnumerable<VideoSocketSettings>?)null,
            vbssSocketSettings: null,
            dataSocketSettings: null,
            mediaSessionId: Guid.NewGuid());

        _audioSocket = mediaSession.AudioSocket;
        _audioSocket.AudioMediaReceived += OnAudioMediaReceived;

        _logger.LogInformation(
            "Media session initialized with unmixed audio. Each frame resolves SSRC (sourceId) at ingestion before routing.");
        return mediaSession;
    }

    private void OnAudioMediaReceived(object? sender, AudioMediaReceivedEventArgs args)
    {
        _ = ProcessAudioMediaReceivedAsync(args);
    }

    private async Task ProcessAudioMediaReceivedAsync(AudioMediaReceivedEventArgs args)
    {
        try
        {
            var buffer = args.Buffer;
            if (buffer is null)
            {
                return;
            }

            var unmixed = buffer.UnmixedAudioBuffers;
            if (unmixed is null || !unmixed.Any())
            {
                if (Interlocked.Increment(ref _loggedNoUnmixed) == 1)
                {
                    _logger.LogWarning(
                        "No UnmixedAudioBuffers in this frame (common briefly at call start). If this repeats forever, Teams is not sending per-participant audio — check meeting type, tenant media policies, and that ReceiveUnmixedMeetingAudio is enabled. Mixed-only audio cannot be attributed.");
                }

                return;
            }

            foreach (var ub in unmixed)
            {
                if (!UnmixedAudioHelpers.TryGetSsrc(ub, out var ssrc))
                {
                    continue;
                }

                var payload = UnmixedAudioHelpers.CopyPayload(ub.Data, ub.Length);
                if (payload.Length == 0)
                {
                    continue;
                }

                await _participantAudioRouter.HandleAudioAsync(ssrc, payload, ub.OriginalSenderTimestamp);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed handling unmixed participant audio.");
        }
    }
}
