using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using MacroPilot.Core.Models;
using MacroPilot.Core.Platform;

namespace MacroPilot.Core.Recording;

public sealed class GlobalMacroRecorder : IDisposable
{
    private const int StopRecordingVirtualKey = 0x78;

    private readonly RecorderOptions _options;
    private readonly NativeMethods.LowLevelKeyboardProc _keyboardProc;
    private readonly NativeMethods.LowLevelMouseProc _mouseProc;
    private readonly Stopwatch _clock = new();

    private nint _keyboardHook;
    private nint _mouseHook;
    private bool _running;
    private long _lastActionTicks;
    private Point? _lastMouseMove;
    private long _lastMouseMoveTicks;

    public GlobalMacroRecorder(RecorderOptions options)
    {
        _options = options;
        _keyboardProc = KeyboardHookCallback;
        _mouseProc = MouseHookCallback;
    }

    public event EventHandler<MacroAction>? ActionRecorded;

    public event EventHandler? StopRequested;

    public bool IsRunning => _running;

    public void Start()
    {
        if (_running)
        {
            return;
        }

        nint moduleHandle = GetCurrentModuleHandle();
        _keyboardHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_KEYBOARD_LL, _keyboardProc, moduleHandle, 0);
        _mouseHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_MOUSE_LL, _mouseProc, moduleHandle, 0);

        if (_keyboardHook == nint.Zero || _mouseHook == nint.Zero)
        {
            Stop();
            throw new InvalidOperationException("Failed to install global keyboard or mouse hooks.");
        }

        _lastActionTicks = 0;
        _lastMouseMove = null;
        _lastMouseMoveTicks = 0;
        _clock.Restart();
        _running = true;
    }

    public void Stop()
    {
        if (_keyboardHook != nint.Zero)
        {
            NativeMethods.UnhookWindowsHookEx(_keyboardHook);
            _keyboardHook = nint.Zero;
        }

        if (_mouseHook != nint.Zero)
        {
            NativeMethods.UnhookWindowsHookEx(_mouseHook);
            _mouseHook = nint.Zero;
        }

        _clock.Stop();
        _running = false;
    }

    public void Dispose()
    {
        Stop();
    }

    private nint KeyboardHookCallback(int nCode, nint wParam, nint lParam)
    {
        if (nCode == NativeMethods.HC_ACTION && _running)
        {
            NativeMethods.KBDLLHOOKSTRUCT hook = Marshal.PtrToStructure<NativeMethods.KBDLLHOOKSTRUCT>(lParam);

            if ((hook.flags & NativeMethods.LLKHF_INJECTED) == 0)
            {
                int message = unchecked((int)wParam);
                bool isDown = message is NativeMethods.WM_KEYDOWN or NativeMethods.WM_SYSKEYDOWN;
                bool isUp = message is NativeMethods.WM_KEYUP or NativeMethods.WM_SYSKEYUP;
                int virtualKey = checked((int)hook.vkCode);

                if (isDown && virtualKey == StopRecordingVirtualKey)
                {
                    StopRequested?.Invoke(this, EventArgs.Empty);
                    return 1;
                }

                if (isDown || isUp)
                {
                    Record(new MacroAction
                    {
                        Type = isDown ? MacroActionType.KeyDown : MacroActionType.KeyUp,
                        VirtualKey = virtualKey,
                        KeyName = NativeMethods.GetKeyName(virtualKey, hook.scanCode, (hook.flags & 1) != 0)
                    });
                }
            }
        }

        return NativeMethods.CallNextHookEx(_keyboardHook, nCode, wParam, lParam);
    }

    private nint MouseHookCallback(int nCode, nint wParam, nint lParam)
    {
        if (nCode == NativeMethods.HC_ACTION && _running)
        {
            NativeMethods.MSLLHOOKSTRUCT hook = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
            if ((hook.flags & NativeMethods.LLMHF_INJECTED) == 0)
            {
                int message = unchecked((int)wParam);
                Point point = new(hook.pt.X, hook.pt.Y);

                switch (message)
                {
                    case NativeMethods.WM_MOUSEMOVE:
                        RecordMouseMove(point);
                        break;
                    case NativeMethods.WM_LBUTTONDOWN:
                        RecordMouseButton(MacroActionType.MouseDown, point, MouseButtonKind.Left);
                        break;
                    case NativeMethods.WM_LBUTTONUP:
                        RecordMouseButton(MacroActionType.MouseUp, point, MouseButtonKind.Left);
                        break;
                    case NativeMethods.WM_RBUTTONDOWN:
                        RecordMouseButton(MacroActionType.MouseDown, point, MouseButtonKind.Right);
                        break;
                    case NativeMethods.WM_RBUTTONUP:
                        RecordMouseButton(MacroActionType.MouseUp, point, MouseButtonKind.Right);
                        break;
                    case NativeMethods.WM_MBUTTONDOWN:
                        RecordMouseButton(MacroActionType.MouseDown, point, MouseButtonKind.Middle);
                        break;
                    case NativeMethods.WM_MBUTTONUP:
                        RecordMouseButton(MacroActionType.MouseUp, point, MouseButtonKind.Middle);
                        break;
                    case NativeMethods.WM_XBUTTONDOWN:
                        RecordMouseButton(MacroActionType.MouseDown, point, GetXButton(hook.mouseData));
                        break;
                    case NativeMethods.WM_XBUTTONUP:
                        RecordMouseButton(MacroActionType.MouseUp, point, GetXButton(hook.mouseData));
                        break;
                    case NativeMethods.WM_MOUSEWHEEL:
                        Record(new MacroAction
                        {
                            Type = MacroActionType.MouseWheel,
                            X = point.X,
                            Y = point.Y,
                            Delta = NativeMethods.HiWord(hook.mouseData)
                        });
                        break;
                }
            }
        }

        return NativeMethods.CallNextHookEx(_mouseHook, nCode, wParam, lParam);
    }

    private void RecordMouseMove(Point point)
    {
        if (!_options.CaptureMouseMoves)
        {
            return;
        }

        long now = _clock.ElapsedTicks;
        int minimumDistance = Math.Max(1, _options.MouseMoveMinimumPixels);
        int minimumInterval = Math.Max(1, _options.MouseMoveMinimumIntervalMs);

        if (_lastMouseMove is { } previous)
        {
            int dx = point.X - previous.X;
            int dy = point.Y - previous.Y;
            long elapsedMs = (now - _lastMouseMoveTicks) * 1000 / Stopwatch.Frequency;
            if ((dx * dx) + (dy * dy) < minimumDistance * minimumDistance && elapsedMs < minimumInterval)
            {
                return;
            }
        }

        _lastMouseMove = point;
        _lastMouseMoveTicks = now;
        Record(new MacroAction
        {
            Type = MacroActionType.MouseMove,
            X = point.X,
            Y = point.Y
        });
    }

    private void RecordMouseButton(MacroActionType type, Point point, MouseButtonKind button)
    {
        Record(new MacroAction
        {
            Type = type,
            X = point.X,
            Y = point.Y,
            Button = button
        });
    }

    private void Record(MacroAction action)
    {
        action.DelayMs = NextDelayMs();
        ActionRecorded?.Invoke(this, action);
    }

    private int NextDelayMs()
    {
        long now = _clock.ElapsedTicks;
        if (_lastActionTicks == 0)
        {
            _lastActionTicks = now;
            return 0;
        }

        long elapsedTicks = now - _lastActionTicks;
        _lastActionTicks = now;
        long elapsedMs = elapsedTicks * 1000 / Stopwatch.Frequency;
        return elapsedMs > int.MaxValue ? int.MaxValue : Math.Max(0, (int)elapsedMs);
    }

    private static MouseButtonKind GetXButton(uint mouseData)
    {
        return NativeMethods.HiWord(mouseData) == 1 ? MouseButtonKind.XButton1 : MouseButtonKind.XButton2;
    }

    private static nint GetCurrentModuleHandle()
    {
        using Process process = Process.GetCurrentProcess();
        string? moduleName = process.MainModule?.ModuleName;
        return NativeMethods.GetModuleHandle(moduleName);
    }
}
