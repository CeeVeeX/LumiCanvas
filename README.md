# LumiCanvas

## NSIS 打包

项目已切换为非 `MSIX` 发布，可使用 `NSIS` 生成安装包。

当前脚本使用 `Release/win-x64` 构建输出目录进行 NSIS 打包（与本地可直接运行目录一致）。

### 前置条件

- 安装 `NSIS`（确保 `makensis.exe` 在 `PATH` 中）
- 安装 `.NET 8 SDK`

### 一键构建安装包

在仓库根目录执行：

`pwsh ./scripts/build-nsis.ps1 -Configuration Release -Runtime win-x64 -Version 1.0.0`

生成文件默认在：

`artifacts/LumiCanvas-Setup-<version>-<runtime>.exe`

### NSIS 脚本位置

`installer/LumiCanvas.nsi`
