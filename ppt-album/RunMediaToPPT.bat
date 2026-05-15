@echo off
chcp 65001 > nul
setlocal

rem 获取当前脚本所在目录
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%GeneratePhotoAlbum.ps1"

rem 如果用户提供了两个参数，则传递；否则直接调用脚本（让脚本自己读 config.ini）
if "%~2"=="" (
    echo 未提供命令行参数，将尝试读取 config.ini 文件。
    powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
) else (
    echo 使用命令行参数：
    echo   媒体文件夹: %~1
    echo   PPT输出路径: %~2
    powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -MediaFolderPath "%~1" -PptFilePath "%~2"
)

pause