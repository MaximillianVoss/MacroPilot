using System.ComponentModel;
using MacroPilot.App.Models;
using MacroPilot.App.Playback;
using MacroPilot.App.Recording;

namespace MacroPilot.App;

public sealed class MainForm : Form
{
    private readonly BindingList<MacroAction> _actions = [];
    private readonly MacroPlayer _player = new();

    private DataGridView _grid = null!;
    private TextBox _nameTextBox = null!;
    private Button _newButton = null!;
    private Button _openButton = null!;
    private Button _saveButton = null!;
    private Button _recordButton = null!;
    private Button _stopButton = null!;
    private Button _playButton = null!;
    private Button _addDelayButton = null!;
    private Button _deleteButton = null!;
    private Button _moveUpButton = null!;
    private Button _moveDownButton = null!;
    private CheckBox _captureMovesCheckBox = null!;
    private CheckBox _moveCursorCheckBox = null!;
    private NumericUpDown _repeatBox = null!;
    private NumericUpDown _speedBox = null!;
    private NumericUpDown _startDelayBox = null!;
    private ToolStripStatusLabel _statusLabel = null!;

    private GlobalMacroRecorder? _recorder;
    private CancellationTokenSource? _playbackCts;
    private string? _currentPath;

    public MainForm()
    {
        UseApplicationIcon();
        BuildUi();
        UpdateInteractionState();
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.F9 && _recorder is not null)
        {
            StopRecording();
            return true;
        }

        if (keyData == Keys.Escape && _playbackCts is not null)
        {
            _playbackCts.Cancel();
            return true;
        }

        return base.ProcessCmdKey(ref msg, keyData);
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        StopRecording();
        _playbackCts?.Cancel();
        base.OnFormClosing(e);
    }

    private void BuildUi()
    {
        Text = "MacroPilot";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(980, 620);
        Size = new Size(1180, 760);

        TableLayoutPanel root = new()
        {
            Dock = DockStyle.Fill,
            RowCount = 3,
            ColumnCount = 1
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));

        root.Controls.Add(BuildTopPanel(), 0, 0);
        root.Controls.Add(BuildGrid(), 0, 1);
        root.Controls.Add(BuildStatusBar(), 0, 2);

        Controls.Add(root);
    }

    private void UseApplicationIcon()
    {
        Icon? icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        if (icon is not null)
        {
            Icon = icon;
        }
    }

    private Control BuildTopPanel()
    {
        TableLayoutPanel panel = new()
        {
            Dock = DockStyle.Fill,
            RowCount = 2,
            ColumnCount = 1,
            Padding = new Padding(10, 8, 10, 4)
        };
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));

        FlowLayoutPanel commands = new()
        {
            Dock = DockStyle.Fill,
            WrapContents = false,
            AutoScroll = true
        };

        _newButton = CreateButton("Новый", NewScript);
        _openButton = CreateButton("Открыть", OpenScriptAsync);
        _saveButton = CreateButton("Сохранить", SaveScriptAsync);
        _recordButton = CreateButton("Запись", StartRecording);
        _stopButton = CreateButton("Стоп", StopActiveWork);
        _playButton = CreateButton("Пуск", PlayScriptAsync);
        _addDelayButton = CreateButton("+ Пауза", AddDelay);
        _deleteButton = CreateButton("Удалить", DeleteSelectedRows);
        _moveUpButton = CreateButton("Выше", MoveSelectedUp);
        _moveDownButton = CreateButton("Ниже", MoveSelectedDown);

        commands.Controls.AddRange(
        [
            _newButton,
            _openButton,
            _saveButton,
            Spacer(12),
            _recordButton,
            _stopButton,
            _playButton,
            Spacer(12),
            _addDelayButton,
            _deleteButton,
            _moveUpButton,
            _moveDownButton
        ]);

        FlowLayoutPanel settings = new()
        {
            Dock = DockStyle.Fill,
            WrapContents = false,
            AutoScroll = true
        };

        _nameTextBox = new TextBox
        {
            Width = 220,
            Text = "Мой макрос"
        };

        _repeatBox = new NumericUpDown
        {
            Minimum = 1,
            Maximum = 999,
            Value = 1,
            Width = 64
        };

        _speedBox = new NumericUpDown
        {
            Minimum = 0.1M,
            Maximum = 10M,
            Increment = 0.1M,
            DecimalPlaces = 1,
            Value = 1M,
            Width = 64
        };

        _startDelayBox = new NumericUpDown
        {
            Minimum = 0M,
            Maximum = 30M,
            Increment = 0.5M,
            DecimalPlaces = 1,
            Value = 1.5M,
            Width = 64
        };

        _captureMovesCheckBox = new CheckBox
        {
            Text = "движения мыши",
            AutoSize = true,
            Checked = false
        };

        _moveCursorCheckBox = new CheckBox
        {
            Text = "ставить курсор",
            AutoSize = true,
            Checked = true
        };

        settings.Controls.AddRange(
        [
            Label("Имя"),
            _nameTextBox,
            Spacer(12),
            Label("Повторы"),
            _repeatBox,
            Spacer(12),
            Label("Скорость"),
            _speedBox,
            Label("x"),
            Spacer(12),
            Label("Старт, сек"),
            _startDelayBox,
            Spacer(12),
            _captureMovesCheckBox,
            Spacer(8),
            _moveCursorCheckBox,
            Spacer(16),
            Label("F9 останавливает запись, Esc останавливает проигрывание")
        ]);

        panel.Controls.Add(commands, 0, 0);
        panel.Controls.Add(settings, 0, 1);
        return panel;
    }

    private Control BuildGrid()
    {
        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            DataSource = _actions,
            AllowUserToAddRows = true,
            AllowUserToDeleteRows = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = true,
            RowHeadersWidth = 48,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = SystemColors.Window,
            BorderStyle = BorderStyle.None
        };

        _grid.Columns.Add(new DataGridViewComboBoxColumn
        {
            HeaderText = "Тип",
            DataPropertyName = nameof(MacroAction.Type),
            DataSource = Enum.GetValues<MacroActionType>(),
            ValueType = typeof(MacroActionType),
            FillWeight = 95
        });
        _grid.Columns.Add(TextColumn("Задержка, мс", nameof(MacroAction.DelayMs), 80));
        _grid.Columns.Add(TextColumn("X", nameof(MacroAction.X), 60));
        _grid.Columns.Add(TextColumn("Y", nameof(MacroAction.Y), 60));
        _grid.Columns.Add(new DataGridViewComboBoxColumn
        {
            HeaderText = "Кнопка",
            DataPropertyName = nameof(MacroAction.Button),
            DataSource = Enum.GetValues<MouseButtonKind>(),
            ValueType = typeof(MouseButtonKind),
            FillWeight = 85
        });
        _grid.Columns.Add(TextColumn("Колесо", nameof(MacroAction.Delta), 70));
        _grid.Columns.Add(TextColumn("VK", nameof(MacroAction.VirtualKey), 70));
        _grid.Columns.Add(TextColumn("Клавиша", nameof(MacroAction.KeyName), 100));
        _grid.Columns.Add(TextColumn("Комментарий", nameof(MacroAction.Comment), 170));

        _grid.DataError += (_, e) =>
        {
            e.ThrowException = false;
            SetStatus("Проверьте значение в таблице: оно не подходит для выбранной колонки.");
        };

        _grid.SelectionChanged += (_, _) => UpdateInteractionState();
        _grid.RowsAdded += (_, _) => UpdateInteractionState();
        _grid.RowsRemoved += (_, _) => UpdateInteractionState();

        return _grid;
    }

    private Control BuildStatusBar()
    {
        StatusStrip statusStrip = new();
        _statusLabel = new ToolStripStatusLabel("Готово");
        statusStrip.Items.Add(_statusLabel);
        return statusStrip;
    }

    private static DataGridViewTextBoxColumn TextColumn(string header, string property, float weight)
    {
        return new DataGridViewTextBoxColumn
        {
            HeaderText = header,
            DataPropertyName = property,
            FillWeight = weight
        };
    }

    private static Button CreateButton(string text, Action action)
    {
        Button button = new()
        {
            Text = text,
            AutoSize = true,
            Height = 34,
            Margin = new Padding(0, 4, 8, 4),
            Padding = new Padding(10, 3, 10, 3)
        };
        button.Click += (_, _) => action();
        return button;
    }

    private static Button CreateButton(string text, Func<Task> action)
    {
        Button button = new()
        {
            Text = text,
            AutoSize = true,
            Height = 34,
            Margin = new Padding(0, 4, 8, 4),
            Padding = new Padding(10, 3, 10, 3)
        };
        button.Click += async (_, _) => await action();
        return button;
    }

    private static Label Label(string text)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            Margin = new Padding(0, 9, 5, 0)
        };
    }

    private static Control Spacer(int width)
    {
        return new Panel
        {
            Width = width,
            Height = 1,
            Margin = Padding.Empty
        };
    }

    private void NewScript()
    {
        if (!ConfirmDiscardIfNeeded())
        {
            return;
        }

        StopActiveWork();
        _currentPath = null;
        _nameTextBox.Text = "Мой макрос";
        _actions.Clear();
        SetStatus("Создан пустой сценарий.");
        UpdateInteractionState();
    }

    private async Task OpenScriptAsync()
    {
        if (!ConfirmDiscardIfNeeded())
        {
            return;
        }

        using OpenFileDialog dialog = new()
        {
            Filter = "MacroPilot script (*.macropilot.json)|*.macropilot.json|JSON (*.json)|*.json|All files (*.*)|*.*",
            Title = "Открыть сценарий"
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            MacroScript script = await ScriptSerializer.LoadAsync(dialog.FileName);
            _currentPath = dialog.FileName;
            _nameTextBox.Text = script.Name;
            _actions.Clear();
            foreach (MacroAction action in script.Actions)
            {
                _actions.Add(action);
            }

            SetStatus($"Открыто действий: {_actions.Count}.");
        }
        catch (Exception ex)
        {
            ShowError("Не удалось открыть сценарий.", ex);
        }

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
            using SaveFileDialog dialog = new()
            {
                Filter = "MacroPilot script (*.macropilot.json)|*.macropilot.json|JSON (*.json)|*.json",
                Title = "Сохранить сценарий",
                FileName = $"{SanitizeFileName(_nameTextBox.Text)}.macropilot.json"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
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

    private void StartRecording()
    {
        if (_playbackCts is not null)
        {
            return;
        }

        if (_actions.Count > 0)
        {
            DialogResult result = MessageBox.Show(
                this,
                "Очистить текущий список действий перед записью?",
                "Запись",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (result == DialogResult.Cancel)
            {
                return;
            }

            if (result == DialogResult.Yes)
            {
                _actions.Clear();
            }
        }

        try
        {
            _recorder = new GlobalMacroRecorder(new RecorderOptions
            {
                CaptureMouseMoves = _captureMovesCheckBox.Checked
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
        SetStatus($"Запись остановлена. Действий: {_actions.Count}.");
        UpdateInteractionState();
    }

    private async Task PlayScriptAsync()
    {
        if (_playbackCts is not null || _recorder is not null || _actions.Count == 0)
        {
            return;
        }

        if (!TryCommitGrid())
        {
            return;
        }

        MacroScript script = BuildScriptFromGrid();
        PlaybackOptions options = new()
        {
            RepeatCount = (int)_repeatBox.Value,
            Speed = (double)_speedBox.Value,
            StartDelayMs = (int)(_startDelayBox.Value * 1000),
            MoveCursorForMouseActions = _moveCursorCheckBox.Checked
        };

        using CancellationTokenSource cts = new();
        _playbackCts = cts;
        UpdateInteractionState();
        SetStatus($"Старт через {options.StartDelayMs / 1000.0:0.0} сек. Esc отменяет проигрывание.");

        Progress<PlaybackProgress> progress = new(p =>
        {
            SetStatus($"Проигрывание: проход {p.RepeatIndex}/{p.RepeatCount}, действие {p.ActionIndex}/{p.ActionCount}.");
        });

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

    private void AddDelay()
    {
        _actions.Add(new MacroAction
        {
            Type = MacroActionType.Delay,
            DelayMs = 1000,
            Comment = "Пауза"
        });
        SetStatus("Добавлена пауза 1000 мс.");
        UpdateInteractionState();
    }

    private void DeleteSelectedRows()
    {
        int[] indexes = SelectedRowIndexes();
        if (indexes.Length == 0)
        {
            return;
        }

        foreach (int index in indexes.OrderDescending())
        {
            if (index >= 0 && index < _actions.Count)
            {
                _actions.RemoveAt(index);
            }
        }

        SetStatus($"Удалено строк: {indexes.Length}.");
        UpdateInteractionState();
    }

    private void MoveSelectedUp()
    {
        int index = CurrentRowIndex();
        if (index <= 0 || index >= _actions.Count)
        {
            return;
        }

        Swap(index, index - 1);
        SelectRow(index - 1);
    }

    private void MoveSelectedDown()
    {
        int index = CurrentRowIndex();
        if (index < 0 || index >= _actions.Count - 1)
        {
            return;
        }

        Swap(index, index + 1);
        SelectRow(index + 1);
    }

    private void Swap(int left, int right)
    {
        (_actions[left], _actions[right]) = (_actions[right], _actions[left]);
        _actions.ResetBindings();
        SetStatus("Порядок действий изменен.");
        UpdateInteractionState();
    }

    private int CurrentRowIndex()
    {
        return _grid.CurrentRow?.Index ?? -1;
    }

    private int[] SelectedRowIndexes()
    {
        if (_grid.SelectedRows.Count > 0)
        {
            return _grid.SelectedRows
                .Cast<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .Select(row => row.Index)
                .Distinct()
                .ToArray();
        }

        return _grid.SelectedCells
            .Cast<DataGridViewCell>()
            .Select(cell => cell.RowIndex)
            .Where(index => index >= 0 && index < _actions.Count)
            .Distinct()
            .ToArray();
    }

    private void SelectRow(int index)
    {
        if (index < 0 || index >= _grid.Rows.Count)
        {
            return;
        }

        _grid.ClearSelection();
        _grid.Rows[index].Selected = true;
        _grid.CurrentCell = _grid.Rows[index].Cells[0];
    }

    private void RecorderOnActionRecorded(object? sender, MacroAction action)
    {
        if (IsDisposed || !IsHandleCreated)
        {
            return;
        }

        BeginInvoke((MethodInvoker)(() =>
        {
            _actions.Add(action);
            SetStatus($"Записано действий: {_actions.Count}. Последнее: {DescribeAction(action)}.");
            UpdateInteractionState();
        }));
    }

    private void RecorderOnStopRequested(object? sender, EventArgs e)
    {
        if (IsDisposed || !IsHandleCreated)
        {
            return;
        }

        BeginInvoke((MethodInvoker)StopRecording);
    }

    private MacroScript BuildScriptFromGrid()
    {
        return new MacroScript
        {
            Name = string.IsNullOrWhiteSpace(_nameTextBox.Text) ? "Untitled macro" : _nameTextBox.Text.Trim(),
            CreatedAt = DateTimeOffset.Now,
            Actions = _actions.ToList()
        };
    }

    private bool TryCommitGrid()
    {
        try
        {
            if (_grid.IsCurrentCellInEditMode)
            {
                _grid.EndEdit();
            }

            BindingContext? bindingContext = BindingContext;
            bindingContext?[_actions].EndCurrentEdit();
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
        if (_actions.Count == 0)
        {
            return true;
        }

        DialogResult result = MessageBox.Show(
            this,
            "Текущий сценарий не будет сохранен автоматически. Продолжить?",
            "MacroPilot",
            MessageBoxButtons.OKCancel,
            MessageBoxIcon.Warning);

        return result == DialogResult.OK;
    }

    private void UpdateInteractionState()
    {
        bool isRecording = _recorder is not null;
        bool isPlaying = _playbackCts is not null;
        bool busy = isRecording || isPlaying;
        bool hasActions = _actions.Count > 0;
        bool hasSelection = SelectedRowIndexes().Length > 0;
        int currentIndex = CurrentRowIndex();

        _newButton.Enabled = !busy;
        _openButton.Enabled = !busy;
        _saveButton.Enabled = !isPlaying && !isRecording;
        _recordButton.Enabled = !busy;
        _stopButton.Enabled = busy;
        _playButton.Enabled = !busy && hasActions;
        _addDelayButton.Enabled = !busy;
        _deleteButton.Enabled = !busy && hasSelection;
        _moveUpButton.Enabled = !busy && currentIndex > 0;
        _moveDownButton.Enabled = !busy && currentIndex >= 0 && currentIndex < _actions.Count - 1;
        _grid.ReadOnly = busy;
        _captureMovesCheckBox.Enabled = !busy;
        _moveCursorCheckBox.Enabled = !busy;
        _repeatBox.Enabled = !busy;
        _speedBox.Enabled = !busy;
        _startDelayBox.Enabled = !busy;
    }

    private void SetStatus(string message)
    {
        _statusLabel.Text = message;
    }

    private void ShowError(string message, Exception ex)
    {
        SetStatus(message);
        MessageBox.Show(this, $"{message}{Environment.NewLine}{ex.Message}", "MacroPilot", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
