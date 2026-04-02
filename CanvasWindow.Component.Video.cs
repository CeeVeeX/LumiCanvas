using System;
using System.IO;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Windows.Media.Core;

namespace LumiCanvas;

public sealed partial class CanvasWindow
{
    private FrameworkElement BuildVideoCard(BoardItemModel item)
    {
        return !string.IsNullOrWhiteSpace(item.SourcePath) && File.Exists(item.SourcePath)
            ? BuildVideoPlayer(item)
            : CreateMissingMediaHint("视频文件不存在或已被移动");
    }

    private FrameworkElement BuildVideoPlayer(BoardItemModel item)
    {
        var mediaPlayer = new MediaPlayerElement
        {
            Stretch = Stretch.Uniform,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            AreTransportControlsEnabled = false,
            Source = MediaSource.CreateFromUri(new Uri(item.SourcePath!))
        };

        mediaPlayer.PointerEntered += VideoPlayer_PointerEntered;
        mediaPlayer.PointerExited += VideoPlayer_PointerExited;
        mediaPlayer.PointerCanceled += VideoPlayer_PointerExited;
        mediaPlayer.PointerCaptureLost += VideoPlayer_PointerCaptureLost;
        return mediaPlayer;
    }
}
