using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Windows.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MacroPilot.Core.Models;
using MacroPilot.Core.Playback;
using MacroPilot.Core.Recording;
using Microsoft.Win32;

namespace MacroPilot.Wpf;

public partial class MainWindow : Window
{
    private readonly MacroPlayer _player = new();

    private GlobalMacroRecorder? _recorder;
    private CancellationTokenSource? _playbackCts;
    private string? _currentPath;
    private bool _refreshQueued;
    private bool _scrollToLastOnRefresh;

    public MainWindow()
    {
        Actions = [];
        InitializeComponent();
        ActionsGrid.ItemsSource = Actions;
        Actions.CollectionChanged += (_, _) =>
        {
            QueueActionsRefresh(scrollToLast: false);
            UpdateInteractionState();
        };
        UpdateDurationControls();
        UpdateInteractionState();
    }

    public ObservableCollection<MacroAction> Actions { get; }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.F9 && _recorder is not null)
        {
            StopRecording();
            e.Handled = true;
        }

        if (e.Key == Key.Escape && _playbackCts is not null)
        {
            _playbackCts.Cancel();
            e.Handled = true;
        }
    }

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        StopRecording();
        _playbackCts?.Cancel();
    }

    private void NewButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ConfirmDiscardIfNeeded())
        {
            return;
        }

        StopActiveWork();
        _currentPath = null;
        NameTextBox.Text = "Мой макрос";
        Actions.Clear();
        QueueActionsRefresh(scrollToLast: false);
        SetStatus("Создан пустой сценарий.");
        UpdateInteractionState();
    }

    private async void OpenButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ConfirmDiscardIfNeeded())
        {
            return;
        }

        OpenFileDialog dialog = new()
        {
            Filter = "MacroPilot script (*.macropilot.json)|*.macropilot.json|JSON (*.json)|*.json|All files (*.*)|*.*",
            Title = "Открыть сценарий"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            MacroScript script = await ScriptSerializer.LoadAsync(dialog.FileName);
            _currentPath = dialog.FileName;
            NameTextBox.Text = script.Name;
            Actions.Clear();
            foreach (MacroAction action in script.Actions)
            {
                Actions.Add(action);
            }

            QueueActionsRefresh(scrollToLast: true);
            SetStatus($"Открыто действий: {Actions.Count}.");
        }
        catch (Exception ex)
        {
            ShowError("Не удалось открыть сценарий.", ex);
        }

        UpdateInteractionState();
    }

    private async void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        await SaveScriptAsync();
    }

    private void RecordButton_Click(object sender, RoutedEventArgs e)
    {
        if (_playbackCts is not null)
        {
            return;
        }

        if (Actions.Count > 0)
        {
            MessageBoxResult result = MessageBox.Show(
                this,
                "Очистить текущий список действий перед записью?",
                "Запись",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel)
            {
                return;
            }

            if (result == MessageBoxResult.Yes)
            {
                Actions.Clear();
            }
        }

        try
        {
            _recorder = new GlobalMacroRecorder(new RecorderOptions
            {
                CaptureMouseMoves = CaptureMovesCheckBox.IsChecked == true
            });
            _recorder.ActionRecorded += RecorderOnActionRecorded;
            _recorder.StopRequested += RecorderOnStopRequested;
            _recorder.Start();
            SetStatus("Идет запись. Нажмите F9 или кнопку Стоп для остановки.");
        }
        catch (Exception ex)
        {
            _recorder?.Dispose();
            _recorder = null;
            ShowError("Не удалось начать запись.", ex);
        }

        UpdateInteractionState();
    }

    private void StopButton_Click(object sender, RoutedEventArgs e)
    {
        StopActiveWork();
    }

    private async void PlayButton_Click(object sender, RoutedEventArgs e)
    {
        await PlayScriptAsync();
    }

    private void AddDelayButton_Click(object sender, RoutedEventArgs e)
    {
        Actions.Add(new MacroAction
        {
            Type = MacroActionType.Delay,
            DelayMs = 1000,
            Comment = "Пауза"
        });
        QueueActionsRefresh(scrollToLast: true);
        SetStatus("Добавлена пауза 1000 мс.");
        UpdateInteractionState();
    }

    private void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
        List<MacroAction> selected = ActionsGrid.SelectedItems
            .OfType<MacroAction>()
            .ToList();

        if (selected.Count == 0)
        {
            return;
        }

        foreach (MacroAction action in selected)
        {
            Actions.Remove(action);
        }

        SetStatus($"Удалено строк: {selected.Count}.");
        QueueActionsRefresh(scrollToLast: false);
        UpdateInteractionState();
    }

    private void MoveUpButton_Click(object sender, RoutedEventArgs e)
    {
        int index = CurrentIndex();
        if (index <= 0)
        {
            return;
        }

        Actions.Move(index, index - 1);
        QueueActionsRefresh(scrollToLast: false);
        SelectRow(index - 1);
        SetStatus("Порядок действий изменен.");
        UpdateInteractionState();
    }

    private void MoveDownButton_Click(object sender, RoutedEventArgs e)
    {
        int index = CurrentIndex();
        if (index < 0 || index >= Actions.Count - 1)
        {
            return;
        }

        Actions.Move(index, index + 1);
        QueueActionsRefresh(scrollToLast: false);
        SelectRow(index + 1);
        SetStatus("Порядок действий изменен.");
        UpdateInteractionState();
    }

    private void RepeatModeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateDurationControls();
    }

    private void ActionsGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateInteractionState();
    }

    private async Task SaveScriptAsync()
    {
        if (!TryCommitGrid())
        {
            return;
        }

        string? path = _currentPath;
        if (string.IsNullOrWhiteSpace(path))
        {
            SaveFileDialog dialog = new()
            {
                Filter = "MacroPilot script (*.macropilot.json)|*.macropilot.json|JSON (*.json)|*.json",
                Title = "Сохранить сценарий",
                FileName = $"{SanitizeFileName(NameTextBox.Text)}.macropilot.json"
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            path = dialog.FileName;
        }

        try
        {
            await ScriptSerializer.SaveAsync(BuildScriptFromGrid(), path);
            _currentPath = path;
            SetStatus($"Сценарий сохранен: {path}");
        }
        catch (Exception ex)
        {
            ShowError("Не удалось сохранить сценарий.", ex);
        }
    }

    private async Task PlayScriptAsync()
    {
        if (_playbackCts is not null || _recorder is not null || Actions.Count == 0)
        {
            return;
        }

        if (!TryCommitGrid())
        {
            return;
        }

        if (!TryBuildPlaybackOptions(out PlaybackOptions? options))
        {
            return;
        }

        MacroScript script = BuildScriptFromGrid();
        using CancellationTokenSource cts = new();
        _playbackCts = cts;
        UpdateInteractionState();
        SetStatus($"Старт через {options!.StartDelayMs / 1000.0:0.0} сек. Esc отменяет проигрывание.");

        Progress<PlaybackProgress> progress = new(UpdatePlaybackProgress);

        try
        {
            await _player.PlayAsync(script, options, progress, cts.Token);
            SetStatus("Проигрывание завершено.");
        }
        catch (OperationCanceledException)
        {
            SetStatus("Проигрывание остановлено.");
        }
        catch (Exception ex)
        {
            ShowError("Ошибка при проигрывании сценария.", ex);
        }
        finally
        {
            _playbackCts = null;
            UpdateInteractionState();
        }
    }

    private void UpdatePlaybackProgress(PlaybackProgress progress)
    {
        string repeatText = progress.IsDurationMode
            ? $"проход {progress.RepeatIndex}, осталось {FormatRemaining(progress.Remaining)}"
            : $"проход {progress.RepeatIndex}/{progress.RepeatCount}";

        SetStatus($"Проигрывание: {repeatText}, действие {progress.ActionIndex}/{progress.ActionCount}.");
    }

    private bool TryBuildPlaybackOptions(out PlaybackOptions? options)
    {
        options = null;

        PlaybackRepeatMode repeatMode = SelectedRepeatMode();
        int repeatCount = 1;
        int durationMinutes = 1;

        if (repeatMode == PlaybackRepeatMode.Count
            && !TryParseInt(RepeatCountTextBox.Text, 1, 999, "Повторы", out repeatCount))
        {
            return false;
        }

        if (repeatMode == PlaybackRepeatMode.Duration
            && !TryParseInt(DurationMinutesTextBox.Text, 1, 1440, "Длительность", out durationMinutes))
        {
            return false;
        }

        if (!TryParseDouble(SpeedTextBox.Text, 0.1, 10.0, "Скорость", out double speed)
            || !TryParseDouble(StartDelayTextBox.Text, 0, 30, "Стартовая задержка", out double startDelaySeconds))
        {
            return false;
        }

        options = new PlaybackOptions
        {
            RepeatMode = repeatMode,
            RepeatCount = repeatCount,
            RepeatDurationMinutes = durationMinutes,
            Speed = speed,
            StartDelayMs = (int)Math.Round(startDelaySeconds * 1000),
            MoveCursorForMouseActions = MoveCursorCheckBox.IsChecked == true
        };
        return true;
    }

    private void StopActiveWork()
    {
        if (_recorder is not null)
        {
            StopRecording();
        }

        _playbackCts?.Cancel();
    }

    private void StopRecording()
    {
        if (_recorder is null)
        {
            return;
        }

        _recorder.ActionRecorded -= RecorderOnActionRecorded;
        _recorder.StopRequested -= RecorderOnStopRequested;
        _recorder.Dispose();
        _recorder = null;
        SetStatus($"Запись остановлена. Действий: {Actions.Count}.");
        UpdateInteractionState();
    }

    private void RecorderOnActionRecorded(object? sender, MacroAction action)
    {
        Dispatcher.Invoke(() =>
        {
            Actions.Add(action);
            QueueActionsRefresh(scrollToLast: true);
            SetStatus($"Записано действий: {Actions.Count}. Последнее: {DescribeAction(action)}.");
            UpdateInteractionState();
        });
    }

    private void RecorderOnStopRequested(object? sender, EventArgs e)
    {
        Dispatcher.Invoke(StopRecording);
    }

    private MacroScript BuildScriptFromGrid()
    {
        return new MacroScript
        {
            Name = string.IsNullOrWhiteSpace(NameTextBox.Text) ? "Untitled macro" : NameTextBox.Text.Trim(),
            CreatedAt = DateTimeOffset.Now,
            Actions = Actions.ToList()
        };
    }

    private bool TryCommitGrid()
    {
        try
        {
            ActionsGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            ActionsGrid.CommitEdit(DataGridEditingUnit.Row, true);
            return true;
        }
        catch (Exception ex)
        {
            ShowError("Не удалось применить правку в таблице.", ex);
            return false;
        }
    }

    private bool ConfirmDiscardIfNeeded()
    {
        if (Actions.Count == 0)
        {
            return true;
        }

        MessageBoxResult result = MessageBox.Show(
            this,
            "Текущий сценарий не будет сохранен автоматически. Продолжить?",
            "MacroPilot",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Warning);

        return result == MessageBoxResult.OK;
    }

    private void UpdateInteractionState()
    {
        DataGrid? actionsGrid = ActionsGrid;
        if (actionsGrid is null)
        {
            return;
        }

        bool isRecording = _recorder is not null;
        bool isPlaying = _playbackCts is not null;
        bool busy = isRecording || isPlaying;
        bool hasActions = Actions.Count > 0;
        bool hasSelection = actionsGrid.SelectedItems.OfType<MacroAction>().Any();
        int currentIndex = CurrentIndex();

        NewButton.IsEnabled = !busy;
        OpenButton.IsEnabled = !busy;
        SaveButton.IsEnabled = !busy;
        RecordButton.IsEnabled = !busy;
        StopButton.IsEnabled = busy;
        PlayButton.IsEnabled = !busy && hasActions;
        AddDelayButton.IsEnabled = !busy;
        DeleteButton.IsEnabled = !busy && hasSelection;
        MoveUpButton.IsEnabled = !busy && currentIndex > 0;
        MoveDownButton.IsEnabled = !busy && currentIndex >= 0 && currentIndex < Actions.Count - 1;
        actionsGrid.IsReadOnly = busy;
        CaptureMovesCheckBox.IsEnabled = !busy;
        MoveCursorCheckBox.IsEnabled = !busy;
        RepeatModeComboBox.IsEnabled = !busy;
        RepeatCountTextBox.IsEnabled = !busy && SelectedRepeatMode() == PlaybackRepeatMode.Count;
        DurationMinutesTextBox.IsEnabled = !busy && SelectedRepeatMode() == PlaybackRepeatMode.Duration;
        SpeedTextBox.IsEnabled = !busy;
        StartDelayTextBox.IsEnabled = !busy;
    }

    private void UpdateDurationControls()
    {
        if (!IsInitialized)
        {
            return;
        }

        RepeatCountTextBox.IsEnabled = SelectedRepeatMode() == PlaybackRepeatMode.Count && _playbackCts is null && _recorder is null;
        DurationMinutesTextBox.IsEnabled = SelectedRepeatMode() == PlaybackRepeatMode.Duration && _playbackCts is null && _recorder is null;
        UpdateInteractionState();
    }

    private PlaybackRepeatMode SelectedRepeatMode()
    {
        return RepeatModeComboBox?.SelectedItem is ComboBoxItem { Tag: PlaybackRepeatMode mode }
            ? mode
            : PlaybackRepeatMode.Count;
    }

    private int CurrentIndex()
    {
        return ActionsGrid?.SelectedItem is MacroAction action ? Actions.IndexOf(action) : -1;
    }

    private void SelectRow(int index)
    {
        if (index < 0 || index >= Actions.Count)
        {
            return;
        }

        ActionsGrid.SelectedItem = Actions[index];
        ActionsGrid.ScrollIntoView(Actions[index]);
    }

    private void SetStatus(string message)
    {
        StatusTextBlock.Text = message;
    }

    private void ShowError(string message, Exception ex)
    {
        SetStatus(message);
        MessageBox.Show(this, $"{message}{Environment.NewLine}{ex.Message}", "MacroPilot", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private bool TryParseInt(string value, int minimum, int maximum, string label, out int result)
    {
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.CurrentCulture, out result)
            && result >= minimum
            && result <= maximum)
        {
            return true;
        }

        ShowError($"{label}: укажите число от {minimum} до {maximum}.", new FormatException("Некорректное значение."));
        return false;
    }

    private bool TryParseDouble(string value, double minimum, double maximum, string label, out double result)
    {
        if (TryParseFlexibleDouble(value, out result)
            && result >= minimum
            && result <= maximum)
        {
            return true;
        }

        ShowError($"{label}: укажите число от {minimum} до {maximum}. Можно использовать точку или запятую.", new FormatException("Некорректное значение."));
        return false;
    }

    private static bool TryParseFlexibleDouble(string value, out double result)
    {
        string trimmed = value.Trim();
        CultureInfo[] cultures =
        [
            CultureInfo.CurrentCulture,
            CultureInfo.InvariantCulture,
            CultureInfo.GetCultureInfo("ru-RU")
        ];

        foreach (CultureInfo culture in cultures)
        {
            if (double.TryParse(trimmed, NumberStyles.Float, culture, out result))
            {
                return true;
            }
        }

        string invariantCandidate = trimmed.Replace(',', '.');
        if (double.TryParse(invariantCandidate, NumberStyles.Float, CultureInfo.InvariantCulture, out result))
        {
            return true;
        }

        string currentSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        string currentCandidate = trimmed.Replace(".", currentSeparator).Replace(",", currentSeparator);
        return double.TryParse(currentCandidate, NumberStyles.Float, CultureInfo.CurrentCulture, out result);
    }

    private void QueueActionsRefresh(bool scrollToLast)
    {
        if (!IsInitialized || ActionsGrid is null)
        {
            return;
        }

        _scrollToLastOnRefresh |= scrollToLast;

        if (_refreshQueued)
        {
            return;
        }

        _refreshQueued = true;
        Dispatcher.BeginInvoke(DispatcherPriority.Background, () =>
        {
            _refreshQueued = false;
            bool shouldScroll = _scrollToLastOnRefresh;
            _scrollToLastOnRefresh = false;

            try
            {
                ActionsGrid.Items.Refresh();
            }
            catch (InvalidOperationException)
            {
                return;
            }

            if (shouldScroll && Actions.Count > 0)
            {
                ActionsGrid.ScrollIntoView(Actions[Actions.Count - 1]);
            }
        });
    }

    private static string FormatRemaining(TimeSpan? remaining)
    {
        if (!remaining.HasValue)
        {
            return "--:--";
        }

        TimeSpan value = remaining.Value;
        return value.TotalHours >= 1
            ? $"{(int)value.TotalHours}:{value.Minutes:00}:{value.Seconds:00}"
            : $"{value.Minutes:00}:{value.Seconds:00}";
    }

    private static string DescribeAction(MacroAction action)
    {
        return action.Type switch
        {
            MacroActionType.MouseDown or MacroActionType.MouseUp => $"{action.Type} {action.Button} ({action.X}, {action.Y})",
            MacroActionType.MouseMove => $"MouseMove ({action.X}, {action.Y})",
            MacroActionType.MouseWheel => $"MouseWheel {action.Delta}",
            MacroActionType.KeyDown or MacroActionType.KeyUp => $"{action.Type} {action.KeyName}",
            MacroActionType.Delay => $"Delay {action.DelayMs} ms",
            _ => action.Type.ToString()
        };
    }

    private static string SanitizeFileName(string value)
    {
        string name = string.IsNullOrWhiteSpace(value) ? "macro" : value.Trim();
        foreach (char invalid in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(invalid, '_');
        }

        return name;
    }
}
