using System.Text.Json.Serialization;

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

    [JsonIgnore]
    public string DelayText => FormatDuration(TimeSpan.FromMilliseconds(Math.Max(0, DelayMs)));

    private static string FormatDuration(TimeSpan value)
    {
        return value.TotalHours >= 1
            ? $"{(int)value.TotalHours}:{value.Minutes:00}:{value.Seconds:00}.{value.Milliseconds / 100}"
            : $"{value.Minutes:00}:{value.Seconds:00}.{value.Milliseconds / 100}";
    }
}
