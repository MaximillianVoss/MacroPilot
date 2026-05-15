using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using MacroPilot.Core.Models;

namespace MacroPilot.Core.Platform;

internal static class NativeMethods
{
    internal const int WH_KEYBOARD_LL = 13;
    internal const int WH_MOUSE_LL = 14;

    internal const int HC_ACTION = 0;

    internal const int WM_KEYDOWN = 0x0100;
    internal const int WM_KEYUP = 0x0101;
    internal const int WM_SYSKEYDOWN = 0x0104;
    internal const int WM_SYSKEYUP = 0x0105;

    internal const int WM_MOUSEMOVE = 0x0200;
    internal const int WM_LBUTTONDOWN = 0x0201;
    internal const int WM_LBUTTONUP = 0x0202;
    internal const int WM_RBUTTONDOWN = 0x0204;
    internal const int WM_RBUTTONUP = 0x0205;
    internal const int WM_MBUTTONDOWN = 0x0207;
    internal const int WM_MBUTTONUP = 0x0208;
    internal const int WM_MOUSEWHEEL = 0x020A;
    internal const int WM_XBUTTONDOWN = 0x020B;
    internal const int WM_XBUTTONUP = 0x020C;

    internal const uint LLKHF_INJECTED = 0x10;
    internal const uint LLMHF_INJECTED = 0x00000001;

    private const uint INPUT_MOUSE = 0;
    private const uint INPUT_KEYBOARD = 1;

    private const uint KEYEVENTF_KEYUP = 0x0002;

    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;
    private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    private const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
    private const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
    private const uint MOUSEEVENTF_XDOWN = 0x0080;
    private const uint MOUSEEVENTF_XUP = 0x0100;
    private const uint MOUSEEVENTF_WHEEL = 0x0800;

    private const uint XBUTTON1 = 0x0001;
    private const uint XBUTTON2 = 0x0002;

    internal delegate nint LowLevelKeyboardProc(int nCode, nint wParam, nint lParam);

    internal delegate nint LowLevelMouseProc(int nCode, nint wParam, nint lParam);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern nint SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, nint hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern nint SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, nint hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool UnhookWindowsHookEx(nint hhk);

    [DllImport("user32.dll")]
    internal static extern nint CallNextHookEx(nint hhk, int nCode, nint wParam, nint lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    internal static extern nint GetModuleHandle(string? lpModuleName);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern uint MapVirtualKey(uint uCode, uint uMapType);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetKeyNameText(int lParam, StringBuilder lpString, int cchSize);

    internal static void MoveCursor(int x, int y)
    {
        if (!SetCursorPos(x, y))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }
    }

    internal static void SendKey(int virtualKey, bool isDown)
    {
        INPUT[] inputs =
        [
            new()
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = checked((ushort)virtualKey),
                        dwFlags = isDown ? 0 : KEYEVENTF_KEYUP
                    }
                }
            }
        ];

        SendInputOrThrow(inputs);
    }

    internal static string GetKeyName(int virtualKey, uint scanCode, bool isExtended)
    {
        uint resolvedScanCode = scanCode == 0 ? MapVirtualKey((uint)virtualKey, 0) : scanCode;
        int lParam = unchecked((int)(resolvedScanCode << 16));
        if (isExtended)
        {
            lParam |= 1 << 24;
        }

        StringBuilder builder = new(64);
        return GetKeyNameText(lParam, builder, builder.Capacity) > 0
            ? builder.ToString()
            : $"VK_{virtualKey}";
    }

    internal static void SendMouseButton(MouseButtonKind button, bool isDown)
    {
        uint flags = button switch
        {
            MouseButtonKind.Left => isDown ? MOUSEEVENTF_LEFTDOWN : MOUSEEVENTF_LEFTUP,
            MouseButtonKind.Right => isDown ? MOUSEEVENTF_RIGHTDOWN : MOUSEEVENTF_RIGHTUP,
            MouseButtonKind.Middle => isDown ? MOUSEEVENTF_MIDDLEDOWN : MOUSEEVENTF_MIDDLEUP,
            MouseButtonKind.XButton1 or MouseButtonKind.XButton2 => isDown ? MOUSEEVENTF_XDOWN : MOUSEEVENTF_XUP,
            _ => 0
        };

        if (flags == 0)
        {
            return;
        }

        uint mouseData = button switch
        {
            MouseButtonKind.XButton1 => XBUTTON1,
            MouseButtonKind.XButton2 => XBUTTON2,
            _ => 0
        };

        SendMouse(flags, mouseData);
    }

    internal static void SendMouseWheel(int delta)
    {
        SendMouse(MOUSEEVENTF_WHEEL, unchecked((uint)delta));
    }

    private static void SendMouse(uint flags, uint mouseData)
    {
        INPUT[] inputs =
        [
            new()
            {
                type = INPUT_MOUSE,
                U = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        mouseData = mouseData,
                        dwFlags = flags
                    }
                }
            }
        ];

        SendInputOrThrow(inputs);
    }

    private static void SendInputOrThrow(INPUT[] inputs)
    {
        uint sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        if (sent != inputs.Length)
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }
    }

    internal static short HiWord(uint value)
    {
        return unchecked((short)((value >> 16) & 0xFFFF));
    }

    [StructLayout(LayoutKind.Sequential)]
    internal readonly struct POINT
    {
        public readonly int X;
        public readonly int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal readonly struct MSLLHOOKSTRUCT
    {
        public readonly POINT pt;
        public readonly uint mouseData;
        public readonly uint flags;
        public readonly uint time;
        public readonly UIntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal readonly struct KBDLLHOOKSTRUCT
    {
        public readonly uint vkCode;
        public readonly uint scanCode;
        public readonly uint flags;
        public readonly uint time;
        public readonly UIntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;
        public InputUnion U;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public MOUSEINPUT mi;

        [FieldOffset(0)]
        public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }
}
