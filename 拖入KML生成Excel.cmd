@echo off
setlocal
chcp 65001 >nul
if "%~1"=="" (
  echo 请把一个 .kml 文件拖到本文件图标上。
  pause
  exit /b 2
)
where py >nul 2>nul
if not errorlevel 1 (
  py -3 "%~dp0kml_to_excel.py" "%~1"
) else (
  where python >nul 2>nul
  if not errorlevel 1 (
    python "%~dp0kml_to_excel.py" "%~1"
  ) else (
    echo 未找到 Python 3。请先在 Windows 安装 Python 3，并勾选 Add Python to PATH。
    pause
    exit /b 3
  )
)
if errorlevel 1 (
  echo 生成失败，请查看上方错误信息。
) else (
  echo 生成完成。Excel 路径见上方。
)
pause
