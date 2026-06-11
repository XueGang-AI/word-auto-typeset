@echo off
setlocal enabledelayedexpansion
cd /d %~dp0

set APP_NAME=Word2PDF
set DIST_DIR=%cd%\dist_release
set BUILD_DIR=%cd%\build
set VENV_DIR=%cd%\.venv-build
set RUNTIME_DIR=%cd%\runtime\libreoffice

if not exist "%RUNTIME_DIR%\program\soffice.exe" (
  echo [ERROR] 未找到便携 LibreOffice:
  echo         %RUNTIME_DIR%\program\soffice.exe
  echo 请先把便携版 LibreOffice 解压到 runtime\libreoffice\ 目录。
  pause
  exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] 当前系统没有可用 python 命令。
  pause
  exit /b 1
)

echo [INFO] 创建打包虚拟环境...
python -m venv "%VENV_DIR%"
if errorlevel 1 (
  echo [ERROR] 创建虚拟环境失败。
  pause
  exit /b 1
)

call "%VENV_DIR%\Scripts\activate.bat"
python -m pip install --upgrade pip >nul
python -m pip install pyinstaller >nul
if errorlevel 1 (
  echo [ERROR] 安装 pyinstaller 失败。
  pause
  exit /b 1
)

if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"

echo [INFO] 开始打包 EXE...
pyinstaller --noconfirm --clean --onefile --name %APP_NAME% --add-data "convert_word_to_pdf.py;." word2pdf_web.py
if errorlevel 1 (
  echo [ERROR] 打包失败。
  pause
  exit /b 1
)

mkdir "%DIST_DIR%"
copy /y "%cd%\dist\%APP_NAME%.exe" "%DIST_DIR%\%APP_NAME%.exe" >nul
mkdir "%DIST_DIR%\runtime"
xcopy /e /i /y "%RUNTIME_DIR%" "%DIST_DIR%\runtime\libreoffice" >nul

echo @echo off> "%DIST_DIR%\启动网页转换.bat"
echo cd /d %%~dp0>> "%DIST_DIR%\启动网页转换.bat"
echo start "" "%APP_NAME%.exe">> "%DIST_DIR%\启动网页转换.bat"

echo Word2PDF 免安装版使用说明> "%DIST_DIR%\使用说明.txt"
echo.>> "%DIST_DIR%\使用说明.txt"
echo 1）双击“启动网页转换.bat”>> "%DIST_DIR%\使用说明.txt"
echo 2）浏览器打开后上传 doc/docx + 名字列表>> "%DIST_DIR%\使用说明.txt"
echo 3）点击开始，下载 zip 结果>> "%DIST_DIR%\使用说明.txt"
echo.>> "%DIST_DIR%\使用说明.txt"
echo 说明：无需安装 Python、Office。>> "%DIST_DIR%\使用说明.txt"

echo [DONE] 免安装交付目录已生成：
echo        %DIST_DIR%

pause
