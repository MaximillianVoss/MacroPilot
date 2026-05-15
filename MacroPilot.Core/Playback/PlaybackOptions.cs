namespace MacroPilot.Core.Playback;

public sealed class PlaybackOptions
{
    public PlaybackRepeatMode RepeatMode { get; init; } = PlaybackRepeatMode.Count;

    public int RepeatCount { get; init; } = 1;

    public int RepeatDurationMinutes { get; init; } = 1;

    public double Speed { get; init; } = 1.0;

    public int StartDelayMs { get; init; } = 1500;

    public bool MoveCursorForMouseActions { get; init; } = true;
}
