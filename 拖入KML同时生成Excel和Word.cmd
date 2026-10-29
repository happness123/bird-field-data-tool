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
  py -3 "%~dp0kml_to_excel_and_word.py" "%~1"
) else (
  where python >nul 2>nul
  if not errorlevel 1 (
    python "%~dp0kml_to_excel_and_word.py" "%~1"
  ) else (
    echo 未找到 Python 3。请先在 Windows 安装 Python 3。
    pause
    exit /b 3
  )
)
if errorlevel 1 (
  echo 生成未全部完成，请查看上方错误信息。
) else (
  echo Excel 和 Word 均已生成，路径见上方。
)
pause
