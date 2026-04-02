using System;
using System.IO;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;

namespace LumiCanvas;

public sealed partial class CanvasWindow
{
    private FrameworkElement BuildImageCard(BoardItemModel item)
    {
        return !string.IsNullOrWhiteSpace(item.SourcePath) && File.Exists(item.SourcePath)
            ? new Image
            {
                Stretch = Stretch.Uniform,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                Source = new BitmapImage(new Uri(item.SourcePath))
            }
            : CreateMissingMediaHint("图片文件不存在或已被移动");
    }
}
