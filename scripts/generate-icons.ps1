param(
    [string]$SourceIcon = "Assets/icon.png",
    [string]$OutputDir = "Assets"
)

# 检查源图标是否存在
if (!(Test-Path $SourceIcon)) {
    Write-Error "Source icon not found: $SourceIcon"
    exit 1
}

# 检查是否安装了 .NET
try {
    $dotnetVersion = dotnet --version
    Write-Host "Using .NET version: $dotnetVersion" -ForegroundColor Green
} catch {
    Write-Error ".NET SDK not found. Please install .NET 8 SDK."
    exit 1
}

Write-Host "Generating icons from $SourceIcon..." -ForegroundColor Cyan

# 创建临时 C# 项目来调用图像处理 API
$tempDir = Join-Path $env:TEMP "LumiCanvasIconGenerator_$(Get-Random)"
New-Item -ItemType Directory -Path $tempDir | Out-Null

# 创建图标生成器项目
$csprojContent = @"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="System.Drawing.Common" Version="8.0.0" />
  </ItemGroup>
</Project>
"@

$programContent = @"
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

class IconGenerator
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: IconGenerator <source> <output-dir>");
            return;
        }

        var sourcePath = args[0];
        var outputDir = args[1];

        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        Directory.CreateDirectory(outputDir);

        using var sourceImage = Image.FromFile(sourcePath);

        // WinUI 3 需要的图标尺寸
        var sizes = new[]
        {
            // 应用程序图标
            (16, "Square44x44Logo.targetsize-16_altform-unplated.png"),
            (24, "Square44x44Logo.targetsize-24_altform-unplated.png"),
            (32, "Square44x44Logo.targetsize-32_altform-unplated.png"),
            (48, "Square44x44Logo.targetsize-48_altform-unplated.png"),
            (64, "Square44x44Logo.targetsize-64_altform-unplated.png"),
            (96, "Square44x44Logo.targetsize-96_altform-unplated.png"),
            (256, "Square44x44Logo.targetsize-256_altform-unplated.png"),

            // 缩放图标 (scale-200 表示 200% DPI)
            (88, "Square44x44Logo.scale-200.png"),  // 44 * 2 = 88
            (300, "Square150x150Logo.scale-200.png"), // 150 * 2 = 300
            (620, "Wide310x150Logo.scale-200.png"), // 宽度 310 * 2 = 620
            (100, "StoreLogo.png"),
            (1240, "SplashScreen.scale-200.png"), // 620 * 2 = 1240
            (48, "LockScreenLogo.scale-200.png")
        };

        foreach (var (size, filename) in sizes)
        {
            var outputPath = Path.Combine(outputDir, filename);

            // 特殊处理宽屏图标
            if (filename.Contains("Wide") || filename.Contains("Splash"))
            {
                var width = size;
                var height = size / 2; // 2:1 比例
                ResizeAndSave(sourceImage, width, height, outputPath);
                Console.WriteLine($"Created: {filename} ({width}x{height})");
            }
            else
            {
                ResizeAndSave(sourceImage, size, size, outputPath);
                Console.WriteLine($"Created: {filename} ({size}x{size})");
            }
        }

        // 生成 .ico 文件用于托盘图标和应用程序图标
        var iconSizes = new[] { 16, 24, 32, 48, 64, 96, 128, 256 };
        var iconPath = Path.Combine(outputDir, "app.ico");
        CreateIcoFile(sourceImage, iconSizes, iconPath);
        Console.WriteLine($"Created: app.ico (multi-size)");

        Console.WriteLine("Icon generation completed successfully!");
    }

    static void ResizeAndSave(Image source, int width, int height, string outputPath)
    {
        using var resized = new Bitmap(width, height);
        using var graphics = Graphics.FromImage(resized);

        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.SmoothingMode = SmoothingMode.HighQuality;
        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        graphics.CompositingQuality = CompositingQuality.HighQuality;

        graphics.Clear(Color.Transparent);
        graphics.DrawImage(source, 0, 0, width, height);

        resized.Save(outputPath, ImageFormat.Png);
    }

    static void CreateIcoFile(Image source, int[] sizes, string outputPath)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms);

        // ICO 文件头
        writer.Write((short)0); // 保留字段
        writer.Write((short)1); // 类型 (1 = ICO)
        writer.Write((short)sizes.Length); // 图标数量

        var images = new List<byte[]>();
        var offset = 6 + (16 * sizes.Length); // 文件头 + 目录条目

        // 写入目录条目
        foreach (var size in sizes)
        {
            using var resized = new Bitmap(size, size);
            using var graphics = Graphics.FromImage(resized);

            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.CompositingQuality = CompositingQuality.HighQuality;

            graphics.Clear(Color.Transparent);
            graphics.DrawImage(source, 0, 0, size, size);

            using var imageMs = new MemoryStream();
            resized.Save(imageMs, ImageFormat.Png);
            var imageData = imageMs.ToArray();
            images.Add(imageData);

            // 写入目录条目
            writer.Write((byte)(size >= 256 ? 0 : size)); // 宽度
            writer.Write((byte)(size >= 256 ? 0 : size)); // 高度
            writer.Write((byte)0); // 调色板颜色数
            writer.Write((byte)0); // 保留
            writer.Write((short)1); // 颜色平面
            writer.Write((short)32); // 每像素位数
            writer.Write(imageData.Length); // 图像数据大小
            writer.Write(offset); // 图像数据偏移
            offset += imageData.Length;
        }

        // 写入图像数据
        foreach (var imageData in images)
        {
            writer.Write(imageData);
        }

        File.WriteAllBytes(outputPath, ms.ToArray());
    }
}
"@

Set-Content -Path (Join-Path $tempDir "IconGenerator.csproj") -Value $csprojContent
Set-Content -Path (Join-Path $tempDir "Program.cs") -Value $programContent

Write-Host "Building icon generator..." -ForegroundColor Yellow

$originalDir = Get-Location
$sourceFullPath = Resolve-Path (Join-Path $originalDir $SourceIcon)
$outputFullPath = Resolve-Path (Join-Path $originalDir $OutputDir)

Push-Location $tempDir
try {
    $buildOutput = dotnet build -c Release -o bin 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) {
        Write-Host $buildOutput -ForegroundColor Red
        Write-Error "Failed to build icon generator"
        exit 1
    }

    Write-Host "Running icon generator..." -ForegroundColor Yellow

    $runOutput = & ".\bin\IconGenerator.exe" $sourceFullPath $outputFullPath 2>&1 | Out-String
    Write-Host $runOutput

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Icon generation failed"
        exit 1
    }
} finally {
    Pop-Location
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`nIcons generated successfully in: $OutputDir" -ForegroundColor Green
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. app.ico has been set as ApplicationIcon in LumiCanvas.csproj" -ForegroundColor White
Write-Host "  2. Tray icon will use app.ico automatically" -ForegroundColor White
Write-Host "  3. Rebuild the project" -ForegroundColor White
