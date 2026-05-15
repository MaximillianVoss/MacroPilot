namespace MacroPilot.Core.Recording;

public sealed class RecorderOptions
{
    public bool CaptureMouseMoves { get; init; }

    public int MouseMoveMinimumPixels { get; init; } = 4;

    public int MouseMoveMinimumIntervalMs { get; init; } = 25;
}
