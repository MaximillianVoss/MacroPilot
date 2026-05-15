using MacroPilot.App.Models;
using MacroPilot.App.Win32;

namespace MacroPilot.App.Playback;

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

        int repeatCount = Math.Max(1, options.RepeatCount);
        double speed = Math.Clamp(options.Speed, 0.1, 10.0);

        for (int repeat = 1; repeat <= repeatCount; repeat++)
        {
            for (int index = 0; index < script.Actions.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                MacroAction action = script.Actions[index];
                int delay = ScaleDelay(action.DelayMs, speed);

                if (delay > 0)
                {
                    await Task.Delay(delay, cancellationToken);
                }

                Execute(action, options);
                progress?.Report(new PlaybackProgress(repeat, repeatCount, index + 1, script.Actions.Count));
            }
        }
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
                NativeMethods.SendKey(action.VirtualKey, isDown: true);
                break;
            case MacroActionType.KeyUp:
                NativeMethods.SendKey(action.VirtualKey, isDown: false);
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
