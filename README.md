# 🧰 Gadget - 个人工具集

这里汇集了我自己编写的各种实用小工具，每个工具独立存放在子目录中，方便使用和维护。

## 📁 目录结构
gadget/
├── README.md # 总览（本文件）
├── ppt-album/ # 工具1：批量媒体转PPT相册
│ ├── GeneratePhotoAlbum.ps1
│ ├── RunMediaToPPT.bat
│ ├── config.ini.example
│ └── README.md
└── (其他工具目录...) # 工具2、工具3...


## 🔧 工具列表

### 1. 批量媒体转PPT相册工具（ppt-album）

将指定文件夹中的所有图片和视频自动转换为PowerPoint幻灯片，每个媒体文件独占一页并等比例铺满全屏（无黑边/白边），视频自动播放。支持追加到已有PPT文件。

- **主要功能**：
  - 支持常见图片格式（jpg, png, bmp, gif）
  - 支持常见视频格式（mp4, avi, mov, wmv, mkv等）
  - 自动识别媒体类型，分别调用正确的插入方式
  - 支持追加模式：多次运行会在同一PPT末尾新增幻灯片
  - 命令行参数与配置文件两种方式指定路径
  - 处理完成后保持PowerPoint打开，便于继续编辑

- **快速使用**：
  1. 进入 `ppt-album/` 目录
  2. 复制 `config.ini.example` 为 `config.ini`，填写媒体文件夹路径和输出PPT路径
  3. 双击 `RunMediaToPPT.bat`，按提示确认即可

- **详细文档**：[ppt-album/README.md](ppt-album/README.md)

---

### 2. （待添加）

未来新工具会在此列出，并附上简要说明和使用链接。

## 📝 使用说明

- 每个工具子目录均包含独立的 `README.md` 和示例配置文件。
- 所有脚本均已测试于 Windows 10/11 + PowerShell 5.1 环境。
- 如遇到执行策略问题，请以管理员身份运行：
  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

## 📄 许可证
  本工具集仅供个人学习、办公自动化使用。

