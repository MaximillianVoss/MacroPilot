# MacroPilot

MacroPilot is a small Windows macro recorder/player written in C# and WinForms.

It focuses on the ReMouse-like core workflow:

- record keyboard key down/up events;
- record mouse clicks, releases, wheel events, and optionally mouse movement;
- replay the recorded script with repeat count, speed multiplier, and start delay;
- save and load editable JSON scripts (`*.macropilot.json`);
- stop recording globally with `F9`;
- cancel playback with `Esc` when the app has focus.

## Requirements

- Windows
- .NET 8 SDK or newer with Windows Desktop runtime

## Build

```powershell
dotnet build MacroPilot.slnx
```

## Run

```powershell
dotnet run --project .\MacroPilot.App\MacroPilot.App.csproj
```

## Usage Notes

1. Click `Запись` to start recording.
2. Perform the actions you want to capture.
3. Press `F9` or click `Стоп`.
4. Review or edit the action table.
5. Set repeat count, speed, and start delay.
6. Click `Пуск`.

Mouse coordinates are stored as absolute screen coordinates. For reliable playback, keep the target windows in the same positions or edit the coordinates in the table.

MacroPilot records only while the visible application is in recording mode. It ignores injected events to avoid recording its own playback output. If the target application runs as administrator, Windows may block input injection unless MacroPilot is also started with equivalent privileges.
