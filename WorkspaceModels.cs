using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace LumiCanvas;

public enum TaskBoardState
{
    Active,
    Completed,
    Abandoned
}

public enum BoardItemKind
{
    Markdown,
    Image,
    Video,
    File,
    TimeTag
}

[Flags]
public enum TimeTagRecurrence
{
    None = 0,
    Monday = 1 << 0,
    Tuesday = 1 << 1,
    Wednesday = 1 << 2,
    Thursday = 1 << 3,
    Friday = 1 << 4,
    Saturday = 1 << 5,
    Sunday = 1 << 6
}

public sealed class WorkspaceSession : INotifyPropertyChanged
{
    private const string StartupBoardTitle = "LumiCanvas 启动板";
    private const string DefaultMarkdownContent = "# 新笔记";
    private const string StartupBoardMarkdownContent = "# 欢迎使用 LumiCanvas\n\n- `Ctrl + Tab` 唤醒桌面侧边栏\n- 点击任务打开独立画布窗口\n- 双击画布空白处创建文本组件\n- 右键空白处添加文件，图片和视频会直接渲染";
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly string _storageFolder;
    private readonly Dictionary<Guid, CancellationTokenSource> _pendingSaves = [];
    private readonly Dictionary<Guid, TaskBoard> _itemOwners = [];
    private readonly Dictionary<Guid, TaskBoard> _deferredSaveTasks = [];
    private TaskBoard? _currentTask;
    private int _deferredSaveDepth;
    private bool _isLoading;

    public WorkspaceSession()
    {
        _storageFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "LumiCanvas",
            "Tasks");

        Directory.CreateDirectory(_storageFolder);
        LoadTasks();

        if (Tasks.Count == 0)
        {
            var board = new TaskBoard(StartupBoardTitle);
            board.Items.Add(new BoardItemModel
            {
                Kind = BoardItemKind.Markdown,
                X = 180,
                Y = 160,
                Width = 380,
                Height = 260,
                Content = StartupBoardMarkdownContent
            });

            RegisterTask(board);
            Tasks.Add(board);
            SaveTaskNow(board);
        }
    }

    public ObservableCollection<TaskBoard> Tasks { get; } = [];

    public string StorageFolderPath => _storageFolder;

    public TaskBoard? CurrentTask
    {
        get => _currentTask;
        set
        {
            if (ReferenceEquals(_currentTask, value))
            {
                return;
            }

            _currentTask = value;
            OnPropertyChanged();
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public TaskBoard AddTask(string title)
    {
        var task = new TaskBoard(title);
        RegisterTask(task);
        Tasks.Add(task);
        CurrentTask = task;
        SaveTask(task);
        return task;
    }

    public void Flush()
    {
        foreach (var task in Tasks)
        {
            SaveTaskNow(task);
        }
    }

    public void BeginDeferredSave()
    {
        _deferredSaveDepth++;
    }

    public void EndDeferredSave()
    {
        if (_deferredSaveDepth == 0)
        {
            return;
        }

        _deferredSaveDepth--;
        if (_deferredSaveDepth != 0)
        {
            return;
        }

        foreach (var task in _deferredSaveTasks.Values.ToList())
        {
            SaveTask(task);
        }

        _deferredSaveTasks.Clear();
    }

    private void LoadTasks()
    {
        _isLoading = true;
        try
        {
            foreach (var file in Directory.EnumerateFiles(_storageFolder, "*.json").OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    var json = File.ReadAllText(file);
                    var document = JsonSerializer.Deserialize<TaskBoardDocument>(json, JsonOptions);
                    if (document is null)
                    {
                        continue;
                    }

                    var wasMigrated = SanitizeLoadedDocument(document);

                    var board = new TaskBoard(document.Id, document.Title)
                    {
                        State = document.State
                    };

                    foreach (var itemDocument in document.Items)
                    {
                        board.Items.Add(new BoardItemModel
                        {
                            Id = itemDocument.Id,
                            Kind = itemDocument.Kind,
                            X = itemDocument.X,
                            Y = itemDocument.Y,
                            Width = itemDocument.Width,
                            Height = itemDocument.Height,
                            ZIndex = itemDocument.ZIndex,
                            Content = itemDocument.Content,
                            SourcePath = itemDocument.SourcePath,
                            TimeTagDueAt = itemDocument.TimeTagDueAt,
                            TimeTagReminderEnabled = itemDocument.TimeTagReminderEnabled,
                            TimeTagRecurrence = itemDocument.TimeTagRecurrence,
                            TimeTagLastReminderAt = itemDocument.TimeTagLastReminderAt,
                            TimeTagMonthlyDays = itemDocument.TimeTagMonthlyDays
                        });
                    }

                    RegisterTask(board);
                    Tasks.Add(board);

                    if (wasMigrated)
                    {
                        SaveTaskNow(board);
                    }
                }
                catch
                {
                }
            }
        }
        finally
        {
            _isLoading = false;
        }
    }

    private void RegisterTask(TaskBoard task)
    {
        task.PropertyChanged += Task_PropertyChanged;
        task.Items.CollectionChanged += TaskItems_CollectionChanged;

        foreach (var item in task.Items)
        {
            RegisterItem(task, item);
        }
    }

    private void RegisterItem(TaskBoard task, BoardItemModel item)
    {
        item.PropertyChanged += Item_PropertyChanged;
        _itemOwners[item.Id] = task;
    }

    private void UnregisterItem(BoardItemModel item)
    {
        item.PropertyChanged -= Item_PropertyChanged;
        _itemOwners.Remove(item.Id);
    }

    private void Task_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_isLoading || sender is not TaskBoard task)
        {
            return;
        }

        QueueTaskSave(task);
    }

    private void TaskItems_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        var task = Tasks.FirstOrDefault(candidate => ReferenceEquals(candidate.Items, sender));
        if (task is null)
        {
            return;
        }

        if (e.OldItems is not null)
        {
            foreach (BoardItemModel item in e.OldItems)
            {
                UnregisterItem(item);
            }
        }

        if (e.NewItems is not null)
        {
            foreach (BoardItemModel item in e.NewItems)
            {
                RegisterItem(task, item);
            }
        }

        if (!_isLoading)
        {
            QueueTaskSave(task);
        }
    }

    private void Item_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_isLoading || sender is not BoardItemModel item)
        {
            return;
        }

        if (_itemOwners.TryGetValue(item.Id, out var task))
        {
            QueueTaskSave(task);
        }
    }

    private void QueueTaskSave(TaskBoard task)
    {
        if (_deferredSaveDepth > 0)
        {
            _deferredSaveTasks[task.Id] = task;
            return;
        }

        SaveTask(task);
    }

    private static bool SanitizeLoadedDocument(TaskBoardDocument document)
    {
        var changed = false;
        var isStartupBoard = string.Equals(document.Title, StartupBoardTitle, StringComparison.Ordinal) ||
                             document.Title.Contains("LumiCanvas", StringComparison.OrdinalIgnoreCase);

        foreach (var item in document.Items)
        {
            if (item.Kind != BoardItemKind.Markdown)
            {
                continue;
            }

            if (isStartupBoard)
            {
                if (NeedsStartupBoardReset(item.Content))
                {
                    item.Content = StartupBoardMarkdownContent;
                    changed = true;
                }

                continue;
            }
        }

        return changed;
    }

    private static bool NeedsStartupBoardReset(string? content)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return true;
        }

        return !string.Equals(content, StartupBoardMarkdownContent, StringComparison.Ordinal);
    }

    private void SaveTask(TaskBoard task)
    {
        lock (_pendingSaves)
        {
            if (_pendingSaves.TryGetValue(task.Id, out var existing))
            {
                existing.Cancel();
                existing.Dispose();
            }

            var cts = new CancellationTokenSource();
            _pendingSaves[task.Id] = cts;
            _ = SaveTaskDelayedAsync(task, cts.Token);
        }
    }

    private async Task SaveTaskDelayedAsync(TaskBoard task, CancellationToken cancellationToken)
    {
        try
        {
            await Task.Delay(250);
            if (cancellationToken.IsCancellationRequested)
            {
                return;
            }

            SaveTaskNow(task);
        }
        finally
        {
            lock (_pendingSaves)
            {
                if (_pendingSaves.TryGetValue(task.Id, out var current) && current.Token == cancellationToken)
                {
                    current.Dispose();
                    _pendingSaves.Remove(task.Id);
                }
            }
        }
    }

    private void SaveTaskNow(TaskBoard task)
    {
        var document = new TaskBoardDocument
        {
            Id = task.Id,
            Title = task.Title,
            State = task.State,
            Items = task.Items.Select(item => new BoardItemDocument
            {
                Id = item.Id,
                Kind = item.Kind,
                X = item.X,
                Y = item.Y,
                Width = item.Width,
                Height = item.Height,
                ZIndex = item.ZIndex,
                Content = item.Kind == BoardItemKind.Markdown ? item.Content : null,
                SourcePath = item.SourcePath,
                TimeTagDueAt = item.Kind == BoardItemKind.TimeTag ? item.TimeTagDueAt : null,
                TimeTagReminderEnabled = item.Kind == BoardItemKind.TimeTag && item.TimeTagReminderEnabled,
                TimeTagRecurrence = item.Kind == BoardItemKind.TimeTag ? item.TimeTagRecurrence : TimeTagRecurrence.None,
                TimeTagLastReminderAt = item.Kind == BoardItemKind.TimeTag ? item.TimeTagLastReminderAt : null,
                TimeTagMonthlyDays = item.Kind == BoardItemKind.TimeTag ? item.TimeTagMonthlyDays : null
            }).ToList()
        };

        var json = JsonSerializer.Serialize(document, JsonOptions);
        var finalPath = Path.Combine(_storageFolder, $"{task.Id}.json");
        var tempPath = finalPath + ".tmp";
        File.WriteAllText(tempPath, json);
        File.Move(tempPath, finalPath, true);
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private sealed class TaskBoardDocument
    {
        public Guid Id { get; set; }

        public string Title { get; set; } = string.Empty;

        public TaskBoardState State { get; set; }

        public List<BoardItemDocument> Items { get; set; } = [];
    }

    private sealed class BoardItemDocument
    {
        public Guid Id { get; set; }

        public BoardItemKind Kind { get; set; }

        public double X { get; set; }

        public double Y { get; set; }

        public double Width { get; set; }

        public double Height { get; set; }

        public int ZIndex { get; set; }

        public string? Content { get; set; }

        public string? SourcePath { get; set; }

        public DateTimeOffset? TimeTagDueAt { get; set; }

        public bool TimeTagReminderEnabled { get; set; }

        public TimeTagRecurrence TimeTagRecurrence { get; set; }

        public DateTimeOffset? TimeTagLastReminderAt { get; set; }

        public string? TimeTagMonthlyDays { get; set; }
    }
}

public sealed class TaskBoard : INotifyPropertyChanged
{
    private string _title;
    private TaskBoardState _state;

    public TaskBoard(string title)
        : this(Guid.NewGuid(), title)
    {
    }

    public TaskBoard(Guid id, string title)
    {
        Id = id;
        _title = title;
        Items.CollectionChanged += OnItemsChanged;
    }

    public Guid Id { get; }

    public string Title
    {
        get => _title;
        set
        {
            if (_title == value)
            {
                return;
            }

            _title = value;
            OnPropertyChanged();
        }
    }

    public TaskBoardState State
    {
        get => _state;
        set
        {
            if (_state == value)
            {
                return;
            }

            _state = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(DisplayStatus));
        }
    }

    public ObservableCollection<BoardItemModel> Items { get; } = [];

    public string DisplayStatus => State switch
    {
        TaskBoardState.Completed => $"已完成 · {Items.Count} 个组件",
        TaskBoardState.Abandoned => $"已放弃 · {Items.Count} 个组件",
        _ => $"进行中 · {Items.Count} 个组件"
    };

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnItemsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        OnPropertyChanged(nameof(DisplayStatus));
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public sealed class BoardItemModel : INotifyPropertyChanged
{
    private BoardItemKind _kind;
    private double _x;
    private double _y;
    private double _width;
    private double _height;
    private int _zIndex;
    private bool _isEditing;
    private string? _content;
    private string? _sourcePath;
    private DateTimeOffset? _timeTagDueAt;
    private bool _timeTagReminderEnabled = true;
    private TimeTagRecurrence _timeTagRecurrence;
    private DateTimeOffset? _timeTagLastReminderAt;
    private string? _timeTagMonthlyDays;

    public Guid Id { get; set; } = Guid.NewGuid();

    public BoardItemKind Kind
    {
        get => _kind;
        set => SetProperty(ref _kind, value);
    }

    public double X
    {
        get => _x;
        set => SetProperty(ref _x, value);
    }

    public double Y
    {
        get => _y;
        set => SetProperty(ref _y, value);
    }

    public double Width
    {
        get => _width;
        set => SetProperty(ref _width, value);
    }

    public double Height
    {
        get => _height;
        set => SetProperty(ref _height, value);
    }

    public int ZIndex
    {
        get => _zIndex;
        set => SetProperty(ref _zIndex, value);
    }

    public bool IsEditing
    {
        get => _isEditing;
        set => SetProperty(ref _isEditing, value);
    }

    public string? Content
    {
        get => _content;
        set => SetProperty(ref _content, value);
    }

    public string? SourcePath
    {
        get => _sourcePath;
        set => SetProperty(ref _sourcePath, value);
    }

    public DateTimeOffset? TimeTagDueAt
    {
        get => _timeTagDueAt;
        set => SetProperty(ref _timeTagDueAt, value);
    }

    public bool TimeTagReminderEnabled
    {
        get => _timeTagReminderEnabled;
        set => SetProperty(ref _timeTagReminderEnabled, value);
    }

    public TimeTagRecurrence TimeTagRecurrence
    {
        get => _timeTagRecurrence;
        set => SetProperty(ref _timeTagRecurrence, value);
    }

    public DateTimeOffset? TimeTagLastReminderAt
    {
        get => _timeTagLastReminderAt;
        set => SetProperty(ref _timeTagLastReminderAt, value);
    }

    public string? TimeTagMonthlyDays
    {
        get => _timeTagMonthlyDays;
        set => SetProperty(ref _timeTagMonthlyDays, value);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void SetProperty<T>(ref T storage, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(storage, value))
        {
            return;
        }

        storage = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
