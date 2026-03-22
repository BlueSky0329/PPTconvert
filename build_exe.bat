@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ========================================
echo  PPTconvert - 打包为独立 exe
echo ========================================
echo.

if not exist ".venv\Scripts\python.exe" (
  echo [1/4] 创建虚拟环境 .venv ...
  python -m venv .venv
  if errorlevel 1 (
    echo 错误：未找到 python。请先安装 Python 3.10+ 并加入 PATH。
    pause
    exit /b 1
  )
) else (
  echo [1/4] 使用已有虚拟环境 .venv
)

echo [2/4] 安装运行依赖与 PyInstaller ...
".venv\Scripts\python.exe" -m pip install -q -U pip
".venv\Scripts\python.exe" -m pip install -q -r requirements.txt -r requirements-build.txt
if errorlevel 1 (
  echo pip 安装失败
  pause
  exit /b 1
)

echo [3/4] 执行 PyInstaller ...
".venv\Scripts\pyinstaller.exe" --noconfirm PPTconvert.spec
if errorlevel 1 (
  echo 打包失败，请查看上方报错
  pause
  exit /b 1
)

echo.
echo [4/4] 完成。
echo  生成文件: dist\PPTconvert.exe
echo  可将 dist\PPTconvert.exe 单独发给他人（对方电脑无需安装 Python）。
echo.
pause
