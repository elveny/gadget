param(
    [string]$MediaFolderPath,
    [string]$PptFilePath
)

# ==================== 读取配置文件函数 ====================
function Read-ConfigFile {
    param([string]$configPath)
    $config = @{}
    if (Test-Path $configPath) {
        Get-Content $configPath | ForEach-Object {
            if ($_ -match '^\s*([^#=]+)\s*=\s*(.*)\s*$') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $config[$key] = $value
            }
        }
    }
    return $config
}

# ==================== 确定路径 ====================
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = Join-Path $scriptDir "config.ini"

# 1. 优先使用命令行参数
if (-not $MediaFolderPath) { $MediaFolderPath = $null }
if (-not $PptFilePath) { $PptFilePath = $null }

# 2. 尝试从 config.ini 读取
$config = Read-ConfigFile $configPath
if (-not $MediaFolderPath -and $config.ContainsKey("MediaFolderPath")) {
    $MediaFolderPath = $config["MediaFolderPath"]
}
if (-not $PptFilePath -and $config.ContainsKey("PptFilePath")) {
    $PptFilePath = $config["PptFilePath"]
}

# 3. 验证路径是否提供
if (-not $MediaFolderPath -or -not $PptFilePath) {
    Write-Host "错误：未提供必要的路径参数！" -ForegroundColor Red
    Write-Host "使用方法：" -ForegroundColor Yellow
    Write-Host "  1. 在命令行中指定： .\GeneratePhotoAlbum.ps1 -MediaFolderPath `"D:\media`" -PptFilePath `"D:\out.pptx`"" -ForegroundColor Yellow
    Write-Host "  2. 或在同级目录下创建 config.ini 文件，内容示例：" -ForegroundColor Yellow
    Write-Host "       MediaFolderPath=D:\media" -ForegroundColor Yellow
    Write-Host "       PptFilePath=D:\out.pptx" -ForegroundColor Yellow
    exit 1
}

# ==================== 用户确认 ====================
Write-Host "`n即将使用以下路径：" -ForegroundColor Cyan
Write-Host "  媒体文件夹: $MediaFolderPath" -ForegroundColor White
Write-Host "  输出PPT文件: $PptFilePath" -ForegroundColor White
Write-Host "`n请确认是否继续？(Y/N) " -ForegroundColor Yellow -NoNewline
$confirm = Read-Host
if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "操作已取消。" -ForegroundColor Red
    exit 0
}

# ==================== 原有脚本逻辑（保持不变） ====================
# 注意：从下面开始，原脚本的变量名需要与 $MediaFolderPath 和 $PptFilePath 保持一致
# 原有脚本中用的是 $mediaFolderPath 和 $pptFilePath，这里已统一

# ---------- 修复中文编码 ----------
chcp 65001 > $null
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ---------- 支持的扩展名 ----------
$imageExtensions = @(".jpg", ".jpeg", ".png", ".bmp", ".gif")
$videoExtensions = @(".mp4", ".avi", ".mov", ".wmv", ".mkv", ".flv", ".m4v")

# ---------- 获取所有媒体文件 ----------
Write-Host "正在扫描文件夹: $MediaFolderPath" -ForegroundColor Cyan
$allFilesRaw = Get-ChildItem -LiteralPath $MediaFolderPath -Recurse -File -ErrorAction SilentlyContinue
if (-not $allFilesRaw) {
    Write-Host "错误：无法读取文件夹或文件夹为空！" -ForegroundColor Red
    exit 1
}

$allFiles = @()
foreach ($file in $allFilesRaw) {
    $ext = $file.Extension.ToLower()
    if ($imageExtensions -contains $ext) {
        $file | Add-Member -MemberType NoteProperty -Name MediaType -Value "Image" -PassThru | Out-Null
        $allFiles += $file
    } elseif ($videoExtensions -contains $ext) {
        $file | Add-Member -MemberType NoteProperty -Name MediaType -Value "Video" -PassThru | Out-Null
        $allFiles += $file
    }
}

if ($allFiles.Count -eq 0) {
    Write-Host "错误：没有找到任何图片或视频！路径: $MediaFolderPath" -ForegroundColor Red
    exit 1
}

$imageCount = ($allFiles | Where-Object { $_.MediaType -eq "Image" }).Count
$videoCount = ($allFiles | Where-Object { $_.MediaType -eq "Video" }).Count
Write-Host "找到 $($allFiles.Count) 个媒体文件（图片: $imageCount, 视频: $videoCount）"

# ---------- 启动或打开 PowerPoint ----------
$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.Visible = -1

if (Test-Path $PptFilePath) {
    Write-Host "PPT 文件已存在，将追加新幻灯片..." -ForegroundColor Cyan
    $presentation = $pptApp.Presentations.Open($PptFilePath)
} else {
    Write-Host "PPT 文件不存在，新建文档..." -ForegroundColor Cyan
    $presentation = $pptApp.Presentations.Add()
}

$slideWidth = $presentation.PageSetup.SlideWidth
$slideHeight = $presentation.PageSetup.SlideHeight
Write-Host "当前幻灯片尺寸：宽=${slideWidth}pt, 高=${slideHeight}pt"

# ---------- 处理媒体文件 ----------
$successCount = 0
foreach ($file in $allFiles) {
    $relativePath = $file.FullName.Substring($MediaFolderPath.Length).TrimStart('\')
    Write-Host "正在处理: $relativePath (类型: $($file.MediaType))"

    $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)

    try {
        if ($file.MediaType -eq "Image") {
            $shape = $slide.Shapes.AddPicture($file.FullName, 0, -1, 0, 0, -1, -1)
        } else {
            $shape = $slide.Shapes.AddMediaObject2($file.FullName, 0, -1, 0, 0, -1, -1)
        }
    } catch {
        Write-Host "  插入失败: $($file.Name) - $_" -ForegroundColor Yellow
        continue
    }

    if ($shape -eq $null -or $shape.Width -eq 0 -or $shape.Height -eq 0) {
        Write-Host "  媒体无效，跳过" -ForegroundColor Yellow
        continue
    }

    $slideRatio = [double]$slideWidth / [double]$slideHeight
    $mediaRatio = [double]$shape.Width / [double]$shape.Height

    if ($mediaRatio -gt $slideRatio) {
        $newHeight = $slideHeight
        $newWidth = $slideHeight * $mediaRatio
        $shape.Height = [double]$newHeight
        $shape.Width  = [double]$newWidth
        $shape.Left   = [double](($slideWidth - $newWidth) / 2)
        $shape.Top    = 0.0
    } else {
        $newWidth = $slideWidth
        $newHeight = $slideWidth / $mediaRatio
        $shape.Width  = [double]$newWidth
        $shape.Height = [double]$newHeight
        $shape.Top    = [double](($slideHeight - $newHeight) / 2)
        $shape.Left   = 0.0
    }

    if ($file.MediaType -eq "Video") {
        try {
            $shape.AnimationSettings.PlaySettings.PlayOnEntry = -1
        } catch {
            Write-Host "  警告：无法设置自动播放" -ForegroundColor Yellow
        }
    }

    $successCount++
}

Write-Host "成功处理 $successCount / $($allFiles.Count) 个媒体文件"

if ($successCount -gt 0) {
    $presentation.SaveAs($PptFilePath)
    Write-Host "PPT 已保存：$PptFilePath" -ForegroundColor Green
    Write-Host "PPT 保持打开状态，您现在可以手动编辑。" -ForegroundColor Cyan
} else {
    Write-Host "没有成功插入任何媒体，不保存" -ForegroundColor Red
}

# 不关闭 PPT，保持打开