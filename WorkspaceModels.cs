using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO.Compression;
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
    private const string ArchiveExtension = ".lumi";
    private const string ArchiveJsonEntryName = "task.json";
    private const string TaskMenuFileName = "task-menu.json";
    private const string StartupBoardTitle = "LumiCanvas 启动板";
    private const string DefaultMarkdownContent = "# 新笔记";
    private const string StartupBoardMarkdownContent = "# 欢迎使用 LumiCanvas\n\n- `Ctrl + Tab` 唤醒桌面侧边栏\n- 点击任务打开独立画布窗口\n- 双击画布空白处创建文本组件\n- 右键空白处添加文件，图片和视频会直接渲染";
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly string _storageFolder;
    private readonly string _taskAssetCacheFolder;
    private readonly string _taskMenuIndexPath;
    private readonly Dictionary<Guid, CancellationTokenSource> _pendingSaves = [];
    private readonly Dictionary<Guid, TaskBoard> _itemOwners = [];
    private readonly Dictionary<Guid, TaskBoard> _deferredSaveTasks = [];
    private readonly Dictionary<Guid, string> _taskArchivePaths = [];
    private readonly Dictionary<string, Guid> _archivePathToTaskId = new(StringComparer.OrdinalIgnoreCase);
    private TaskBoard? _currentTask;
    private int _deferredSaveDepth;
    private bool _isLoading;

    public WorkspaceSession()
    {
        _storageFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "LumiCanvas",
            "Tasks");

        _taskAssetCacheFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "LumiCanvas",
            "TaskAssets");
        _taskMenuIndexPath = Path.Combine(_storageFolder, TaskMenuFileName);

        Directory.CreateDirectory(_storageFolder);
        Directory.CreateDirectory(_taskAssetCacheFolder);
        if (File.Exists(_taskMenuIndexPath))
        {
            TrySetHiddenAttribute(_taskMenuIndexPath);
        }
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

    public string GetTaskAssetsDirectory(Guid taskId)
    {
        var taskRoot = Path.Combine(_taskAssetCacheFolder, taskId.ToString("N"));
        var assetsRoot = Path.Combine(taskRoot, "assets");
        Directory.CreateDirectory(assetsRoot);
        return assetsRoot;
    }

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

        SaveTaskMenuIndex();
    }

    public bool TryEnsureTaskFromArchivePath(string archivePath, out TaskBoard? task)
    {
        task = null;
        if (string.IsNullOrWhiteSpace(archivePath))
        {
            return false;
        }

        string normalizedPath;
        try
        {
            normalizedPath = Path.GetFullPath(archivePath);
        }
        catch
        {
            return false;
        }

        if (!File.Exists(normalizedPath) || !string.Equals(Path.GetExtension(normalizedPath), ArchiveExtension, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (_archivePathToTaskId.TryGetValue(normalizedPath, out var existingTaskId))
        {
            task = Tasks.FirstOrDefault(candidate => candidate.Id == existingTaskId);
            return task is not null;
        }

        try
        {
            var loadedTask = LoadTaskBoardFromArchive(normalizedPath, out var wasMigrated);
            if (loadedTask is null)
            {
                return false;
            }

            loadedTask = EnsureUniqueTaskIdentity(loadedTask, normalizedPath);

            RegisterTask(loadedTask);
            Tasks.Add(loadedTask);
            SetArchivePathMapping(loadedTask.Id, normalizedPath);
            SaveTaskMenuIndex();

            if (wasMigrated)
            {
                SaveTaskNow(loadedTask);
            }

            task = loadedTask;
            return true;
        }
        catch (Exception ex)
        {
            App.WriteDiagnostic($"WorkspaceSession.TryEnsureTaskFromArchivePath: {archivePath}", ex);
            return false;
        }
    }

    public string? GetTaskArchivePath(Guid taskId)
    {
        return _taskArchivePaths.TryGetValue(taskId, out var path) ? path : null;
    }

    public bool RemoveTaskFromMenu(Guid taskId)
    {
        var task = Tasks.FirstOrDefault(candidate => candidate.Id == taskId);
        if (task is null)
        {
            return false;
        }

        task.PropertyChanged -= Task_PropertyChanged;
        task.Items.CollectionChanged -= TaskItems_CollectionChanged;
        foreach (var item in task.Items.ToList())
        {
            UnregisterItem(item);
        }

        Tasks.Remove(task);

        if (CurrentTask?.Id == taskId)
        {
            CurrentTask = null;
        }

        if (_taskArchivePaths.TryGetValue(taskId, out var archivePath))
        {
            _archivePathToTaskId.Remove(archivePath);
        }

        _taskArchivePaths.Remove(taskId);
        _deferredSaveTasks.Remove(taskId);
        SaveTaskMenuIndex();
        return true;
    }

    public int PruneMissingMenuTasks()
    {
        var missingTaskIds = _taskArchivePaths
            .Where(entry => string.IsNullOrWhiteSpace(entry.Value) || !File.Exists(entry.Value))
            .Select(entry => entry.Key)
            .ToList();

        if (missingTaskIds.Count == 0)
        {
            return 0;
        }

        var removed = 0;
        foreach (var taskId in missingTaskIds)
        {
            if (RemoveTaskFromMenu(taskId))
            {
                removed++;
            }
        }

        return removed;
    }

    public void ClearTaskAssetsCache()
    {
        if (!Directory.Exists(_taskAssetCacheFolder))
        {
            return;
        }

        try
        {
            Directory.Delete(_taskAssetCacheFolder, recursive: true);
        }
        catch
        {
            return;
        }

        Directory.CreateDirectory(_taskAssetCacheFolder);
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
            var candidateArchives = Directory.EnumerateFiles(_storageFolder, $"*{ArchiveExtension}")
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase);

            foreach (var file in candidateArchives)
            {
                try
                {
                    if (!File.Exists(file))
                    {
                        continue;
                    }

                    var board = LoadTaskBoardFromArchive(file, out var wasMigrated);
                    if (board is null)
                    {
                        continue;
                    }

                    board = EnsureUniqueTaskIdentity(board, file);

                    RegisterTask(board);
                    Tasks.Add(board);
                    SetArchivePathMapping(board.Id, file);

                    if (wasMigrated)
                    {
                        SaveTaskNow(board);
                    }
                }
                catch (Exception ex)
                {
                    App.WriteDiagnostic($"WorkspaceSession.LoadTasks: {file}", ex);
                }
            }
        }
        finally
        {
            _isLoading = false;
        }
    }

    private TaskBoard EnsureUniqueTaskIdentity(TaskBoard task, string archivePath)
    {
        if (!Tasks.Any(existing => existing.Id == task.Id))
        {
            return task;
        }

        var replacement = new TaskBoard(Guid.NewGuid(), task.Title)
        {
            State = task.State
        };

        foreach (var item in task.Items)
        {
            replacement.Items.Add(new BoardItemModel
            {
                Id = Guid.NewGuid(),
                Kind = item.Kind,
                X = item.X,
                Y = item.Y,
                Width = item.Width,
                Height = item.Height,
                ZIndex = item.ZIndex,
                Content = item.Content,
                SourcePath = item.SourcePath,
                TimeTagDueAt = item.TimeTagDueAt,
                TimeTagReminderEnabled = item.TimeTagReminderEnabled,
                TimeTagRecurrence = item.TimeTagRecurrence,
                TimeTagLastReminderAt = item.TimeTagLastReminderAt,
                TimeTagMonthlyDays = item.TimeTagMonthlyDays
            });
        }

        App.WriteDiagnostic("WorkspaceSession.EnsureUniqueTaskIdentity", new InvalidOperationException($"Duplicate task id detected for archive: {archivePath}"));
        return replacement;
    }

    private TaskBoard? LoadTaskBoardFromArchive(string archivePath, out bool wasMigrated)
    {
        wasMigrated = false;
        using var zip = ZipFile.OpenRead(archivePath);
        var jsonEntry = zip.GetEntry(ArchiveJsonEntryName);
        if (jsonEntry is null)
        {
            return null;
        }

        using var jsonStream = jsonEntry.Open();
        using var reader = new StreamReader(jsonStream);
        var json = reader.ReadToEnd();
        var document = JsonSerializer.Deserialize<TaskBoardDocument>(json, JsonOptions);
        if (document is null)
        {
            return null;
        }

        wasMigrated = SanitizeLoadedDocument(document);

        var taskRoot = Path.Combine(_taskAssetCacheFolder, document.Id.ToString("N"));
        var taskAssetRoot = Path.Combine(taskRoot, "assets");
        Directory.CreateDirectory(taskAssetRoot);
        foreach (var entry in zip.Entries.Where(entry => entry.FullName.StartsWith("assets/", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(entry.Name)))
        {
            var relativePath = entry.FullName.Replace('/', Path.DirectorySeparatorChar);
            var targetPath = Path.Combine(taskRoot, relativePath);
            var targetDirectory = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrWhiteSpace(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            entry.ExtractToFile(targetPath, overwrite: true);
        }

        return CreateTaskBoardFromDocument(document, ResolveArchiveSourcePath);
    }

    private TaskBoard CreateTaskBoardFromDocument(TaskBoardDocument document, Func<Guid, string?, string?> sourcePathResolver)
    {
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
                SourcePath = sourcePathResolver(document.Id, itemDocument.SourcePath),
                TimeTagDueAt = itemDocument.TimeTagDueAt,
                TimeTagReminderEnabled = itemDocument.TimeTagReminderEnabled,
                TimeTagRecurrence = itemDocument.TimeTagRecurrence,
                TimeTagLastReminderAt = itemDocument.TimeTagLastReminderAt,
                TimeTagMonthlyDays = itemDocument.TimeTagMonthlyDays
            });
        }

        return board;
    }

    private string? ResolveArchiveSourcePath(Guid taskId, string? sourcePath)
    {
        if (string.IsNullOrWhiteSpace(sourcePath) || Path.IsPathRooted(sourcePath))
        {
            return sourcePath;
        }

        return Path.Combine(_taskAssetCacheFolder, taskId.ToString("N"), sourcePath.Replace('/', Path.DirectorySeparatorChar));
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
                SourcePath = NormalizeSourcePathForArchive(task.Id, item.SourcePath),
                TimeTagDueAt = item.Kind == BoardItemKind.TimeTag ? item.TimeTagDueAt : null,
                TimeTagReminderEnabled = item.Kind == BoardItemKind.TimeTag && item.TimeTagReminderEnabled,
                TimeTagRecurrence = item.Kind == BoardItemKind.TimeTag ? item.TimeTagRecurrence : TimeTagRecurrence.None,
                TimeTagLastReminderAt = item.Kind == BoardItemKind.TimeTag ? item.TimeTagLastReminderAt : null,
                TimeTagMonthlyDays = item.Kind == BoardItemKind.TimeTag ? item.TimeTagMonthlyDays : null
            }).ToList()
        };

        var json = JsonSerializer.Serialize(document, JsonOptions);
        var finalPath = ResolveArchivePath(task);
        var tempPath = finalPath + ".tmp";
        if (File.Exists(tempPath))
        {
            File.Delete(tempPath);
        }

        var previousPath = _taskArchivePaths.GetValueOrDefault(task.Id);

        using (var archive = ZipFile.Open(tempPath, ZipArchiveMode.Create))
        {
            var jsonEntry = archive.CreateEntry(ArchiveJsonEntryName, CompressionLevel.Optimal);
            using (var entryStream = jsonEntry.Open())
            using (var writer = new StreamWriter(entryStream))
            {
                writer.Write(json);
            }

            var taskRoot = Path.Combine(_taskAssetCacheFolder, task.Id.ToString("N"));
            var assetsRoot = Path.Combine(taskRoot, "assets");
            if (Directory.Exists(assetsRoot))
            {
                foreach (var file in Directory.EnumerateFiles(assetsRoot, "*", SearchOption.AllDirectories))
                {
                    var relativeEntry = Path.GetRelativePath(taskRoot, file).Replace(Path.DirectorySeparatorChar, '/');
                    archive.CreateEntryFromFile(file, relativeEntry, CompressionLevel.Optimal);
                }
            }
        }

        File.Move(tempPath, finalPath, true);
        SetArchivePathMapping(task.Id, finalPath);
        SaveTaskMenuIndex();

        if (!string.IsNullOrWhiteSpace(previousPath) &&
            !string.Equals(previousPath, finalPath, StringComparison.OrdinalIgnoreCase) &&
            File.Exists(previousPath))
        {
            try
            {
                File.Delete(previousPath);
            }
            catch
            {
            }
        }

        var legacyJsonPath = Path.Combine(_storageFolder, $"{task.Id}.json");
        if (File.Exists(legacyJsonPath))
        {
            try
            {
                File.Delete(legacyJsonPath);
            }
            catch
            {
            }
        }
    }

    private string? NormalizeSourcePathForArchive(Guid taskId, string? sourcePath)
    {
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            return sourcePath;
        }

        var taskRoot = Path.Combine(_taskAssetCacheFolder, taskId.ToString("N"));
        if (!Path.IsPathRooted(sourcePath) || !sourcePath.StartsWith(taskRoot, StringComparison.OrdinalIgnoreCase))
        {
            return sourcePath;
        }

        return Path.GetRelativePath(taskRoot, sourcePath).Replace(Path.DirectorySeparatorChar, '/');
    }

    private string ResolveArchivePath(TaskBoard task)
    {
        if (_taskArchivePaths.TryGetValue(task.Id, out var existingPath) &&
            !string.IsNullOrWhiteSpace(existingPath) &&
            !IsUnderStorageFolder(existingPath))
        {
            return existingPath;
        }

        var safeName = BuildSafeFileName(task.Title);
        var candidatePath = Path.Combine(_storageFolder, safeName + ArchiveExtension);

        if (_taskArchivePaths.TryGetValue(task.Id, out var existingPathForCandidate) &&
            string.Equals(existingPathForCandidate, candidatePath, StringComparison.OrdinalIgnoreCase))
        {
            return candidatePath;
        }

        if (!File.Exists(candidatePath) ||
            (_taskArchivePaths.TryGetValue(task.Id, out existingPath) && string.Equals(existingPath, candidatePath, StringComparison.OrdinalIgnoreCase)))
        {
            return candidatePath;
        }

        for (var index = 2; ; index++)
        {
            var conflictCandidate = Path.Combine(_storageFolder, $"{safeName} ({index}){ArchiveExtension}");
            if (!File.Exists(conflictCandidate) ||
                (_taskArchivePaths.TryGetValue(task.Id, out existingPath) && string.Equals(existingPath, conflictCandidate, StringComparison.OrdinalIgnoreCase)))
            {
                return conflictCandidate;
            }
        }
    }

    private static string BuildSafeFileName(string? rawName)
    {
        var name = string.IsNullOrWhiteSpace(rawName) ? "未命名任务" : rawName.Trim();
        var invalidChars = Path.GetInvalidFileNameChars();
        var safeChars = name.Select(ch => invalidChars.Contains(ch) ? '_' : ch).ToArray();
        var safeName = new string(safeChars).Trim().TrimEnd('.');
        return string.IsNullOrWhiteSpace(safeName) ? "未命名任务" : safeName;
    }

    private bool IsUnderStorageFolder(string path)
    {
        return path.StartsWith(_storageFolder, StringComparison.OrdinalIgnoreCase);
    }

    private void SetArchivePathMapping(Guid taskId, string archivePath)
    {
        var normalized = Path.GetFullPath(archivePath);
        if (_taskArchivePaths.TryGetValue(taskId, out var previous))
        {
            _archivePathToTaskId.Remove(previous);
        }

        _taskArchivePaths[taskId] = normalized;
        _archivePathToTaskId[normalized] = taskId;
    }

    private List<string> LoadTaskMenuIndex()
    {
        try
        {
            if (!File.Exists(_taskMenuIndexPath))
            {
                return [];
            }

            var json = File.ReadAllText(_taskMenuIndexPath);
            var index = JsonSerializer.Deserialize<TaskMenuDocument>(json, JsonOptions);
            return index?.ArchivePaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => Path.GetFullPath(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? [];
        }
        catch (Exception ex)
        {
            App.WriteDiagnostic("WorkspaceSession.LoadTaskMenuIndex", ex);
            return [];
        }
    }

    private void SaveTaskMenuIndex()
    {
        try
        {
            var index = new TaskMenuDocument
            {
                ArchivePaths = _taskArchivePaths.Values
                    .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                    .ToList()
            };

            var json = JsonSerializer.Serialize(index, JsonOptions);
            File.WriteAllText(_taskMenuIndexPath, json);
            TrySetHiddenAttribute(_taskMenuIndexPath);
        }
        catch (Exception ex)
        {
            App.WriteDiagnostic("WorkspaceSession.SaveTaskMenuIndex", ex);
        }
    }

    private static void TrySetHiddenAttribute(string filePath)
    {
        try
        {
            var current = File.GetAttributes(filePath);
            if (!current.HasFlag(FileAttributes.Hidden))
            {
                File.SetAttributes(filePath, current | FileAttributes.Hidden);
            }
        }
        catch
        {
        }
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private sealed class TaskMenuDocument
    {
        public List<string> ArchivePaths { get; set; } = [];
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
