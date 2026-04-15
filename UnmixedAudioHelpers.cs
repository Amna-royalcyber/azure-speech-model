using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Skype.Bots.Media;

namespace TeamsMediaBot;

/// <summary>
/// SSRC for Teams unmixed audio is the media stream <c>sourceId</c> on each <see cref="UnmixedAudioBuffer"/> (SDK may expose it as SourceId / StreamSourceId).
/// </summary>
internal static class UnmixedAudioHelpers
{
    public static bool TryGetSsrc(UnmixedAudioBuffer ub, out uint ssrc)
    {
        ssrc = 0;
        var none = (uint)DominantSpeakerChangedEventArgs.None;
        try
        {
            foreach (var propName in new[] { "SourceId", "StreamSourceId", "MediaSourceId" })
            {
                var p = ub.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (p is null)
                {
                    continue;
                }

                var val = p.GetValue(ub);
                switch (val)
                {
                    case uint u when u != 0 && u != none:
                        ssrc = u;
                        return true;
                    case int i when i > 0:
                        ssrc = (uint)i;
                        return true;
                }
            }
        }
        catch
        {
            // fall through
        }

        var active = Convert.ToUInt32(ub.ActiveSpeakerId);
        if (active != 0 && active != none)
        {
            ssrc = active;
            return true;
        }

        return false;
    }

    public static byte[] CopyPayload(IntPtr ptr, long length)
    {
        if (ptr == IntPtr.Zero || length <= 0 || length > int.MaxValue)
        {
            return Array.Empty<byte>();
        }

        var bytes = new byte[(int)length];
        Marshal.Copy(ptr, bytes, 0, (int)length);
        return bytes;
    }
}
