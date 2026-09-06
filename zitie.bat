@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

rem ===== 找 Python（优先 python，其次 py 启动器）=====
set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY (
  where py >nul 2>nul && set "PY=py"
)
if not defined PY (
  echo [错误] 未找到 Python。请先安装 Python 3.9+：
  echo        https://www.python.org/downloads/
  echo 安装时务必勾选 "Add Python to PATH"，然后重新运行本脚本。
  echo.
  pause
  exit /b 1
)

rem ===== 安装依赖 python-docx（仅首次）=====
%PY% -c "import docx" >nul 2>nul
if errorlevel 1 (
  echo 首次运行，正在安装依赖 python-docx ...
  %PY% -m pip install -r requirements.txt
)

rem ===== 首次运行生成 .env 配置（可用记事本修改）=====
if not exist ".env" (
  if exist ".env.example" copy ".env.example" ".env" >nul
  echo 已根据模板生成 .env 配置文件，可用记事本打开修改字体 / 标题 / 排版。
  echo.
)

rem ===== 生成字帖 =====
rem 用法：
rem   双击本文件                → 按 .env 配置生成 content.txt
rem   把 .txt 拖到本文件上      → 生成该文件的字帖
rem   命令行: zitie.bat 内容.txt --font wenkai --pdf
if "%~1"=="" (
  echo 未传入内容文件，按 .env 配置生成 content.txt 的字帖 ...
  %PY% zitie.py content.txt
) else (
  %PY% zitie.py %*
)

echo.
echo 完成。生成的 .docx / .pdf 在本目录下，可用 Word/WPS 打开打印。
pause
