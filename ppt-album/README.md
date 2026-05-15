# 批量媒体文件转 PPT 工具

将指定文件夹中的所有图片和视频自动转换为 PowerPoint 幻灯片，每个媒体文件独占一页，并自动铺满全屏（等比例缩放、居中、超出部分裁剪）。支持追加到已有 PPT 文件。

## 功能特性

- 支持常见图片格式（.jpg, .jpeg, .png, .bmp, .gif）
- 支持常见视频格式（.mp4, .avi, .mov, .wmv, .mkv, .flv, .m4v）
- 自动识别文件类型，分别调用正确的插入方式
- 每个媒体文件占一页空白幻灯片，并填充整个页面（无黑边/白边）
- 视频自动播放（进入页面即开始）
- 支持追加模式：如果 PPT 已存在，则在末尾添加新幻灯片
- 支持命令行参数和配置文件两种方式指定路径
- 处理完成后保持 PowerPoint 打开，方便继续编辑
- 支持中文路径（需正确编码保存文件）

## 环境要求

- Windows 操作系统
- Microsoft PowerPoint（2007 及以上版本，推荐 2016+）
- PowerShell 5.1 或更高版本（Windows 10/11 自带）

## 文件说明

将以下三个文件放在同一个文件夹中：

| 文件名 | 说明 |
|--------|------|
| GeneratePhotoAlbum.ps1 | 主 PowerShell 脚本 |
| RunMediaToPPT.bat | 批处理启动器（可选，用于双击运行） |
| config.ini | 配置文件（可选，用于默认路径） |

## 快速开始

### 方法一：使用配置文件（推荐）

1. 编辑 config.ini，填入你的媒体文件夹路径和输出 PPT 路径：
   MediaFolderPath=D:\我的媒体文件夹
   PptFilePath=D:\生成的相册.pptx

2. 双击 RunMediaToPPT.bat

3. 脚本会显示读取到的路径，输入 Y 确认后开始处理

### 方法二：使用命令行参数（优先级最高）

打开命令提示符（cmd）或 PowerShell，进入脚本目录，执行：

RunMediaToPPT.bat "D:\媒体文件夹" "D:\输出.pptx"

或者直接调用 PowerShell：

powershell -ExecutionPolicy Bypass -File GeneratePhotoAlbum.ps1 -MediaFolderPath "D:\媒体文件夹" -PptFilePath "D:\输出.pptx"

### 方法三：仅配置文件，直接调用 PowerShell

右键 GeneratePhotoAlbum.ps1 -> “使用 PowerShell 运行”，脚本会自动读取同目录下的 config.ini。

## 配置详解

### config.ini 示例

# 这是注释
MediaFolderPath = D:\doc\给家人的PPT\成长日历\玥儿成长日历\resources
PptFilePath = D:\doc\给家人的PPT\成长日历\script\output.pptx

- 键名区分大小写，必须为 MediaFolderPath 和 PptFilePath
- 等号两边可以有空格
- 路径可以包含中文，但 config.ini 文件必须保存为 UTF-8 with BOM 编码（见注意事项）
- 路径中的反斜杠 \ 或正斜杠 / 均可

### 命令行参数

| 参数名 | 说明 |
|--------|------|
| -MediaFolderPath | 存放图片和视频的文件夹路径 |
| -PptFilePath | 生成的 PPT 文件完整路径（.pptx） |

## 优先级规则

1. 命令行参数（如果提供）
2. config.ini（如果命令行未提供）
3. 如果两者均未提供，脚本报错并提示

## 使用示例

### 示例 1：首次生成相册

config.ini：
MediaFolderPath=C:\Users\Administrator\Desktop\毕业旅行照片视频
PptFilePath=C:\Users\Administrator\Desktop\毕业旅行.pptx

双击 RunMediaToPPT.bat，输入 Y 后，脚本会新建 毕业旅行.pptx，将文件夹中所有图片和视频按名称排序插入。

### 示例 2：追加新内容到已有 PPT

原有 PPT 包含 10 张幻灯片，现在 resources 文件夹新增了 5 个视频。再次运行脚本（使用相同输出路径），会在 PPT 末尾追加 5 张新幻灯片，原有内容保留。

### 示例 3：只处理图片（排除视频）

修改 GeneratePhotoAlbum.ps1 中的 $videoExtensions 为空数组：
$videoExtensions = @()

## 注意事项

### 文件编码

- GeneratePhotoAlbum.ps1 必须保存为 UTF-8 with BOM 编码，否则中文路径会乱码。
- config.ini 同样推荐 UTF-8 with BOM。如果使用普通 UTF-8（无 BOM），请手动修改脚本中的 Get-Content 行为：
  Get-Content $configPath -Encoding UTF8

### 视频格式兼容性

- 强烈推荐使用 H.264 + AAC 编码的 .mp4 文件，PowerPoint 兼容性最好。
- 某些 .avi、.mov 或特殊编码的 .mp4 可能无法插入或播放。

### 性能与文件大小

- 每个视频都会被嵌入到 PPT 中，导致文件体积迅速增大（一个 100MB 的视频会使 PPT 增加约 100MB）。
- 插入大量高清视频时，脚本执行时间较长，请耐心等待。

### PowerPoint 进程

- 脚本处理完成后不会关闭 PowerPoint，方便您立即检查或继续编辑。
- 如果多次运行脚本，每次都会打开新的 PPT 窗口（除非手动关闭之前的）。建议运行前关闭已打开的未保存 PPT。

### 路径中的空格

- 如果路径包含空格，在命令行中使用双引号包裹，例如：
  RunMediaToPPT.bat "D:\My Photos" "D:\My Album.pptx"

### 执行策略

- 如果系统禁止运行 PowerShell 脚本，请以管理员身份运行一次：
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

- 批处理文件已包含 -ExecutionPolicy Bypass，无需手动设置。

## 常见问题

### Q1：视频插入失败，提示 E_FAIL

A：通常是因为视频编码不被 PowerPoint 支持。解决方案：
- 使用格式工厂、HandBrake 等工具将视频转换为 H.264 + AAC 的 .mp4
- 或直接在 PPT 中手动插入（插入 -> 视频 -> 此设备）

### Q2：中文路径显示乱码，找不到文件夹

A：请确保所有脚本和配置文件均保存为 UTF-8 with BOM 编码。用记事本打开，另存为，编码选择 UTF-8（记事本的 UTF-8 就是带 BOM 的）。

### Q3：视频不能自动播放

A：检查视频形状的动画设置。脚本已尝试设置 PlayOnEntry = -1，但某些情况下可能失效。可以手动在 PPT 中设置：选中视频 -> “播放”选项卡 -> “开始”选择“自动”。

### Q4：PPT 文件被占用，无法保存

A：请先关闭正在使用该 PPT 文件的 PowerPoint 窗口，再运行脚本。

### Q5：如何让图片和视频按修改时间排序？

A：修改脚本中的 Sort-Object Name 为 Sort-Object LastWriteTime。

### Q6：只处理当前目录，不递归子文件夹？

A：删除脚本中 Get-ChildItem 的 -Recurse 参数。

## 文件清单

项目文件夹/
├── GeneratePhotoAlbum.ps1   # 主脚本（必须）
├── RunMediaToPPT.bat        # 批处理启动器（可选）
├── config.ini               # 配置文件（可选）
└── README.md                # 本文件

## 许可证

本工具仅供个人学习、办公自动化使用。

## 更新日志

- v1.0（2025-05-15）
  - 初始版本
  - 支持图片和视频批量插入
  - 支持追加模式和配置文件
  - 保持 PPT 打开

祝您使用愉快！如有任何问题，欢迎反馈。