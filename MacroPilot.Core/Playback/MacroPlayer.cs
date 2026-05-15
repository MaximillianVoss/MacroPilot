using MacroPilot.Core.Models;
using MacroPilot.Core.Platform;

namespace MacroPilot.Core.Playback;

public sealed class MacroPlayer
{
    public async Task PlayAsync(
        MacroScript script,
        PlaybackOptions options,
        IProgress<PlaybackProgress>? progress,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(script);
        ArgumentNullException.ThrowIfNull(options);

        if (options.StartDelayMs > 0)
        {
            await Task.Delay(options.StartDelayMs, cancellationToken);
        }

        double speed = Math.Clamp(options.Speed, 0.1, 10.0);
        int repeatCount = Math.Max(1, options.RepeatCount);
        DateTimeOffset? stopAt = options.RepeatMode == PlaybackRepeatMode.Duration
            ? DateTimeOffset.UtcNow.AddMinutes(Math.Max(1, options.RepeatDurationMinutes))
            : null;

        for (int repeat = 1; ShouldStartRepeat(repeat, repeatCount, stopAt); repeat++)
        {
            for (int index = 0; index < script.Actions.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                MacroAction action = script.Actions[index];
                int delay = ScaleDelay(action.DelayMs, speed);

                if (!await WaitForActionDelayAsync(delay, stopAt, cancellationToken))
                {
                    return;
                }

                Execute(action, options);
                progress?.Report(new PlaybackProgress(
                    repeat,
                    stopAt.HasValue ? 0 : repeatCount,
                    index + 1,
                    script.Actions.Count)
                {
                    Remaining = RemainingUntil(stopAt)
                });
            }
        }
    }

    private static bool ShouldStartRepeat(int repeat, int repeatCount, DateTimeOffset? stopAt)
    {
        if (!stopAt.HasValue)
        {
            return repeat <= repeatCount;
        }

        return DateTimeOffset.UtcNow < stopAt.Value;
    }

    private static async Task<bool> WaitForActionDelayAsync(
        int delayMs,
        DateTimeOffset? stopAt,
        CancellationToken cancellationToken)
    {
        if (delayMs <= 0)
        {
            return !stopAt.HasValue || DateTimeOffset.UtcNow < stopAt.Value;
        }

        if (!stopAt.HasValue)
        {
            await Task.Delay(delayMs, cancellationToken);
            return true;
        }

        TimeSpan remaining = stopAt.Value - DateTimeOffset.UtcNow;
        if (remaining <= TimeSpan.Zero)
        {
            return false;
        }

        TimeSpan requestedDelay = TimeSpan.FromMilliseconds(delayMs);
        if (remaining < requestedDelay)
        {
            await Task.Delay(remaining, cancellationToken);
            return false;
        }

        await Task.Delay(requestedDelay, cancellationToken);
        return true;
    }

    private static TimeSpan? RemainingUntil(DateTimeOffset? stopAt)
    {
        if (!stopAt.HasValue)
        {
            return null;
        }

        TimeSpan remaining = stopAt.Value - DateTimeOffset.UtcNow;
        return remaining > TimeSpan.Zero ? remaining : TimeSpan.Zero;
    }

    private static int ScaleDelay(int delayMs, double speed)
    {
        if (delayMs <= 0)
        {
            return 0;
        }

        double scaled = delayMs / speed;
        return scaled > int.MaxValue ? int.MaxValue : Math.Max(0, (int)Math.Round(scaled));
    }

    private static void Execute(MacroAction action, PlaybackOptions options)
    {
        switch (action.Type)
        {
            case MacroActionType.Delay:
                break;
            case MacroActionType.MouseMove:
                if (options.MoveCursorForMouseActions)
                {
                    NativeMethods.MoveCursor(action.X, action.Y);
                }
                break;
            case MacroActionType.MouseDown:
                MoveCursorIfNeeded(action, options);
                NativeMethods.SendMouseButton(action.Button, isDown: true);
                break;
            case MacroActionType.MouseUp:
                MoveCursorIfNeeded(action, options);
                NativeMethods.SendMouseButton(action.Button, isDown: false);
                break;
            case MacroActionType.MouseWheel:
                MoveCursorIfNeeded(action, options);
                NativeMethods.SendMouseWheel(action.Delta);
                break;
            case MacroActionType.KeyDown:
                NativeMethods.SendKey(action.VirtualKey, action.ScanCode, action.IsExtendedKey, isDown: true);
                break;
            case MacroActionType.KeyUp:
                NativeMethods.SendKey(action.VirtualKey, action.ScanCode, action.IsExtendedKey, isDown: false);
                break;
            default:
                throw new NotSupportedException($"Unsupported macro action type: {action.Type}");
        }
    }

    private static void MoveCursorIfNeeded(MacroAction action, PlaybackOptions options)
    {
        if (options.MoveCursorForMouseActions)
        {
            NativeMethods.MoveCursor(action.X, action.Y);
        }
    }
}
