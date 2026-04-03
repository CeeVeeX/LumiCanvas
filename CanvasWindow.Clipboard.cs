using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage;
using Windows.Storage.Streams;

namespace LumiCanvas;

public sealed partial class CanvasWindow
{
    private sealed class ClipboardBoardItemPayload
    {
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

    private void InitializeClipboardIntegration()
    {
        Clipboard.ContentChanged += Clipboard_ContentChanged;
    }

    private void UninitializeClipboardIntegration()
    {
        Clipboard.ContentChanged -= Clipboard_ContentChanged;
    }

    private void Clipboard_ContentChanged(object? sender, object e)
    {
        _copiedBoardItems.Clear();
    }

    private Task<bool> CopySelectedItemsToClipboardAsync()
    {
        if (_session.CurrentTask is null || _selectedItemIds.Count == 0)
        {
            return Task.FromResult(false);
        }

        var selectedItems = _session.CurrentTask.Items
            .Where(item => _selectedItemIds.Contains(item.Id))
            .OrderBy(item => item.ZIndex)
            .ToList();

        if (selectedItems.Count == 0)
        {
            return Task.FromResult(false);
        }

        _copiedBoardItems = selectedItems.Select(item => new ClipboardBoardItemPayload
        {
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
        }).ToList();

        return Task.FromResult(_copiedBoardItems.Count > 0);
    }

    private async Task<bool> PasteCopiedItemsAsync(Windows.Foundation.Point position)
    {
        if (_session.CurrentTask is null || _copiedBoardItems.Count == 0)
        {
            return false;
        }

        var minX = _copiedBoardItems.Min(item => item.X);
        var minY = _copiedBoardItems.Min(item => item.Y);
        var pastedItems = new List<BoardItemModel>(_copiedBoardItems.Count);

        _session.BeginDeferredSave();
        try
        {
            foreach (var item in _copiedBoardItems)
            {
                var pastedItem = new BoardItemModel
                {
                    Kind = item.Kind,
                    X = position.X + (item.X - minX),
                    Y = position.Y + (item.Y - minY),
                    Width = item.Width,
                    Height = item.Height,
                    ZIndex = _highestZIndex + 1,
                    Content = item.Content,
                    SourcePath = await CloneSourcePathForPasteAsync(item.SourcePath),
                    TimeTagDueAt = item.TimeTagDueAt,
                    TimeTagReminderEnabled = item.TimeTagReminderEnabled,
                    TimeTagRecurrence = item.TimeTagRecurrence,
                    TimeTagLastReminderAt = item.TimeTagLastReminderAt,
                    TimeTagMonthlyDays = item.TimeTagMonthlyDays
                };

                _session.CurrentTask.Items.Add(pastedItem);
                AddBoardItemView(pastedItem);
                pastedItems.Add(pastedItem);
            }
        }
        finally
        {
            _session.EndDeferredSave();
        }

        SetSelectedItems(pastedItems);
        return pastedItems.Count > 0;
    }

    private async Task<bool> PasteFromClipboardAsync()
    {
        if (_session.CurrentTask is null)
        {
            return false;
        }

        var dataPackage = Clipboard.GetContent();
        var position = GetPreferredPasteWorldPoint();
        var hasChanges = false;

        if (await PasteCopiedItemsAsync(position))
        {
            return true;
        }

        if (dataPackage.Contains(StandardDataFormats.StorageItems))
        {
            var items = await dataPackage.GetStorageItemsAsync();
            foreach (var file in items.OfType<StorageFile>())
            {
                if (string.IsNullOrWhiteSpace(file.Path))
                {
                    continue;
                }

                var kind = ResolveFileKind(file.Path);
                if (kind is not (BoardItemKind.Image or BoardItemKind.Video))
                {
                    continue;
                }

                var storedPath = await CopyClipboardFileToTaskAssetsAsync(file);
                if (string.IsNullOrWhiteSpace(storedPath))
                {
                    continue;
                }

                AddPastedMediaItem(kind, storedPath, position);
                position = new Windows.Foundation.Point(position.X + 28, position.Y + 28);
                hasChanges = true;
            }
        }

        if (!hasChanges && dataPackage.Contains(StandardDataFormats.Bitmap))
        {
            var bitmap = await dataPackage.GetBitmapAsync();
            var storedPath = await SaveClipboardBitmapAsync(bitmap);
            if (!string.IsNullOrWhiteSpace(storedPath))
            {
                AddPastedMediaItem(BoardItemKind.Image, storedPath, position);
                hasChanges = true;
            }
        }

        if (!hasChanges && dataPackage.Contains(StandardDataFormats.Text))
        {
            var text = await dataPackage.GetTextAsync();
            if (!string.IsNullOrWhiteSpace(text))
            {
                var markdownItem = CreateMarkdownItem(position);
                markdownItem.Content = text;
                _session.CurrentTask.Items.Add(markdownItem);
                AddBoardItemView(markdownItem);
                hasChanges = true;
            }
        }

        return hasChanges;
    }

    private async Task<string?> CloneSourcePathForPasteAsync(string? sourcePath)
    {
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
        {
            return sourcePath;
        }

        var extension = Path.GetExtension(sourcePath);
        var targetPath = BuildTaskAssetPath(extension);
        await Task.Run(() => File.Copy(sourcePath, targetPath, overwrite: true));
        return targetPath;
    }

    private void AddPastedMediaItem(BoardItemKind kind, string sourcePath, Windows.Foundation.Point position)
    {
        if (_session.CurrentTask is null)
        {
            return;
        }

        var item = new BoardItemModel
        {
            Kind = kind,
            X = position.X,
            Y = position.Y,
            Width = kind == BoardItemKind.Image ? 360 : 420,
            Height = kind == BoardItemKind.Image ? 240 : 300,
            ZIndex = _highestZIndex + 1,
            SourcePath = sourcePath
        };

        _session.CurrentTask.Items.Add(item);
        AddBoardItemView(item);
    }

    private async Task<string?> CopyClipboardFileToTaskAssetsAsync(StorageFile sourceFile)
    {
        if (_session.CurrentTask is null)
        {
            return null;
        }

        var extension = Path.GetExtension(sourceFile.Path);
        var targetPath = BuildTaskAssetPath(extension);
        await Task.Run(() => File.Copy(sourceFile.Path, targetPath, overwrite: true));
        return targetPath;
    }

    private async Task<string?> SaveClipboardBitmapAsync(RandomAccessStreamReference bitmapReference)
    {
        if (_session.CurrentTask is null)
        {
            return null;
        }

        var targetPath = BuildTaskAssetPath(".png");
        using var stream = await bitmapReference.OpenReadAsync();
        await using var sourceStream = stream.AsStreamForRead();
        await using var targetStream = File.Create(targetPath);
        await sourceStream.CopyToAsync(targetStream);
        return targetPath;
    }

    private string BuildTaskAssetPath(string extension)
    {
        var assetFolder = _session.GetTaskAssetsDirectory(_session.CurrentTask!.Id);
        var safeExtension = string.IsNullOrWhiteSpace(extension) ? ".bin" : extension;
        var fileName = $"{DateTime.UtcNow:yyyyMMddHHmmssfff}_{Guid.NewGuid():N}{safeExtension}";
        return Path.Combine(assetFolder, fileName);
    }
}
