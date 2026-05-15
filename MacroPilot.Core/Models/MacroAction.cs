namespace MacroPilot.Core.Models;

public sealed class MacroAction
{
    public MacroActionType Type { get; set; }

    public int DelayMs { get; set; }

    public int X { get; set; }

    public int Y { get; set; }

    public MouseButtonKind Button { get; set; }

    public int Delta { get; set; }

    public int VirtualKey { get; set; }

    public int ScanCode { get; set; }

    public bool IsExtendedKey { get; set; }

    public string KeyName { get; set; } = string.Empty;

    public string Comment { get; set; } = string.Empty;
}
