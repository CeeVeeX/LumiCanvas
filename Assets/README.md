# 图标使用说明

## 源图标

源图标文件：`Assets/icon.png` (1490x1490 像素)

## 生成的图标

运行 `scripts/generate-icons.ps1` 会从源图标生成以下文件：

### 应用程序图标
- **app.ico** - 多尺寸 .ico 文件 (16, 24, 32, 48, 64, 96, 128, 256 像素)
  - 用作程序图标（显示在任务栏、文件资源管理器等）
  - 用作托盘图标

### WinUI 3 资源图标
- **Square44x44Logo.targetsize-*.png** - 应用程序徽标 (16, 24, 32, 48, 64, 96, 256 像素)
- **Square44x44Logo.scale-200.png** - 88x88 (44 * 200% DPI)
- **Square150x150Logo.scale-200.png** - 300x300 (150 * 200% DPI)
- **Wide310x150Logo.scale-200.png** - 620x310 (310x155 * 200% DPI)
- **StoreLogo.png** - 100x100
- **SplashScreen.scale-200.png** - 1240x620 (620x310 * 200% DPI)
- **LockScreenLogo.scale-200.png** - 48x48

## 如何更新图标

1. 将新的图标文件（PNG 格式，推荐 1024x1024 或更大）放置在 `Assets/icon.png`
2. 运行生成脚本：
   ```powershell
   pwsh -ExecutionPolicy Bypass -File .\scripts\generate-icons.ps1
   ```
3. 重新构建项目

## 项目配置

### LumiCanvas.csproj
- `<ApplicationIcon>Assets\app.ico</ApplicationIcon>` - 设置程序图标
- 所有生成的图标文件都已添加到 `<Content>` 项中

### MainWindow.xaml.cs
- 托盘图标加载逻辑已更新为使用 `Assets/app.ico`
- 如果找不到自定义图标，会回退到系统默认图标

## 注意事项

- 确保源图标是正方形（1:1 比例）以获得最佳效果
- 对于宽屏图标（Wide、Splash），脚本会自动调整为 2:1 比例
- 生成的 .ico 文件包含多个尺寸，系统会根据需要自动选择合适的尺寸
