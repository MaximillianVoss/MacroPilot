# MacroPilot

MacroPilot is a small Windows macro recorder/player written in C#.

It focuses on the ReMouse-like core workflow:

- record keyboard key down/up events;
- record mouse clicks, releases, wheel events, and optionally mouse movement;
- replay the recorded script with repeat count, speed multiplier, and start delay;
- replay the recorded script for a configured duration in minutes in the WPF app;
- show the script duration, planned run time, and remaining playback time;
- save and load editable JSON scripts (`*.macropilot.json`);
- start recording in the WPF app with `Ctrl+F9`;
- stop recording globally with `Shift+F9`;
- cancel playback with `Esc` when the app has focus.

## Projects

- `MacroPilot.Core` - shared script model, JSON serialization, global recorder, and playback engine.
- `MacroPilot.App` - original Windows Forms shell.
- `MacroPilot.Wpf` - newer WPF shell with a cleaner layout and duration-based replay mode.

## Requirements

- Windows
- .NET 8 SDK or newer with Windows Desktop runtime

## Build

```powershell
dotnet build MacroPilot.slnx
```

## Run

```powershell
dotnet run --project .\MacroPilot.Wpf\MacroPilot.Wpf.csproj
```

To run the Windows Forms version:

```powershell
dotnet run --project .\MacroPilot.App\MacroPilot.App.csproj
```

## Usage Notes

1. Click `Запись` or press `Ctrl+F9` in the WPF app to start recording.
2. Perform the actions you want to capture.
3. Press `Shift+F9` or click `Стоп`.
4. Review or edit the action table. The `Время` column shows each action delay in a readable format.
5. Set repeat count or switch the WPF app to minute-based duration mode. The unused repeat/duration field is hidden automatically.
6. Click `Пуск`.

Mouse coordinates are stored as absolute screen coordinates. For reliable playback, keep the target windows in the same positions or edit the coordinates in the table.

MacroPilot records only while the visible application is in recording mode. It ignores injected events to avoid recording its own playback output. If the target application runs as administrator, Windows may block input injection unless MacroPilot is also started with equivalent privileges.
