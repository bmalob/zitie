#!/usr/bin/env bash
# 字帖生成器 · macOS / Linux 一键脚本
# 用法：
#   ./zitie.sh                          → 按 .env 配置生成 content.txt 的字帖
#   ./zitie.sh 内容.txt                  → 生成指定文件的字帖
#   ./zitie.sh 内容.txt --font wenkai --pdf   → 追加任意命令行参数
#   也可把 .txt 文件拖到终端的本脚本上运行
set -euo pipefail
cd "$(dirname "$0")"

# ===== 找 Python（优先 python3）=====
PY=""
for c in python3 python; do
  if command -v "$c" >/dev/null 2>&1; then PY="$c"; break; fi
done
if [ -z "$PY" ]; then
  echo "[错误] 未找到 Python。请先安装 Python 3.9+："
  echo "  macOS：  brew install python   （或 https://www.python.org/downloads/）"
  echo "  Linux：  sudo apt install python3 python3-pip"
  exit 1
fi

# ===== 安装依赖 python-docx（仅首次）=====
if ! "$PY" -c "import docx" >/dev/null 2>&1; then
  echo "首次运行，正在安装依赖 python-docx ..."
  "$PY" -m pip install -r requirements.txt
fi

# ===== 首次运行生成 .env 配置（可编辑）=====
if [ ! -f ".env" ] && [ -f ".env.example" ]; then
  cp ".env.example" ".env"
  echo "已根据模板生成 .env 配置文件，可编辑修改字体 / 标题 / 排版。"
  echo
fi

# ===== 生成字帖 =====
if [ $# -eq 0 ]; then
  echo "未传入内容文件，按 .env 配置生成 content.txt 的字帖 ..."
  "$PY" zitie.py content.txt
else
  "$PY" zitie.py "$@"
fi

echo
echo "完成。生成的 .docx / .pdf 在本目录下，可用 Word / WPS / LibreOffice 打开打印。"
