namespace MacroPilot.App.Playback;

public sealed class PlaybackOptions
{
    public int RepeatCount { get; init; } = 1;

    public double Speed { get; init; } = 1.0;

    public int StartDelayMs { get; init; } = 1500;

    public bool MoveCursorForMouseActions { get; init; } = true;
}
