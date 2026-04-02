using System;
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
    private async Task PasteFromClipboardAsync()
    {
        if (_session.CurrentTask is null)
        {
            return;
        }

        var dataPackage = Clipboard.GetContent();
        var position = GetViewportCenterWorldPoint();
        var hasChanges = false;

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
