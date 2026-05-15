namespace MacroPilot.Core.Playback;

public sealed record PlaybackProgress(int RepeatIndex, int RepeatCount, int ActionIndex, int ActionCount)
{
    public TimeSpan? Remaining { get; init; }

    public bool IsDurationMode => RepeatCount <= 0;
}
