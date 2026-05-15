namespace MacroPilot.App.Models;

public sealed class MacroScript
{
    public string Name { get; set; } = "Untitled macro";

    public DateTimeOffset CreatedAt { get; set; } = DateTimeOffset.Now;

    public List<MacroAction> Actions { get; set; } = [];
}
