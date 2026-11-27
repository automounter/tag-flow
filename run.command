#!/bin/bash
# macOS で Finder からダブルクリックして起動する
cd "$(dirname "$0")"
if [ -x .venv/bin/python ]; then
  exec .venv/bin/python main.py
elif [ -x runtime/python/bin/python3.12 ]; then
  exec runtime/python/bin/python3.12 main.py
else
  exec python3 main.py
fi
