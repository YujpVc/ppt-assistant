#!/usr/bin/env bash
# Run a PPT-engine Python script with an interpreter that actually has the
# deck dependencies. DSH's non-login shell may resolve `python3` to a runtime
# without python-pptx; this picks the first usable interpreter instead.
#
#   bash <engine-root>/run_py.sh build_deck.py spec.json out.pptx
#   PPT_PYTHON=/path/to/python bash <engine-root>/run_py.sh quality_check.py deck.pptx
set -euo pipefail

if [[ $# -lt 1 ]]; then
  echo "run_py: usage: run_py.sh <script.py> [args...]" >&2
  exit 64
fi

# DSH's Tauri shell may inject PYTHONHOME/PYTHONPATH for its own runtime;
# those poison standalone Python/conda installs before encodings can load.
unset PYTHONHOME PYTHONPATH

engine_root=$(CDPATH= cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
script=$1
shift
[[ -f "$script" || -f "$engine_root/$script" ]] || { echo "run_py: script not found: $script" >&2; exit 66; }

candidates=()
[[ -n "${PPT_PYTHON:-}" ]] && candidates+=("$PPT_PYTHON")
candidates+=(
  "${CONDA_PREFIX:-}/bin/python"
  "$HOME/anaconda3/bin/python"
  "$HOME/miniconda3/bin/python"
  "python3"
  "python3.13"
  "python3.12"
  "python3.11"
  "python3.10"
  "python3.9"
)

for python in "${candidates[@]}"; do
  [[ -n "$python" ]] || continue
  command -v "$python" >/dev/null 2>&1 || continue
  if "$python" -c 'import pptx, PIL' >/dev/null 2>&1; then
    exec "$python" "$script" "$@"
  fi
done

echo "run_py: 没有找到带 python-pptx/Pillow 的解释器。请安装 requirements.txt，或设置 PPT_PYTHON。" >&2
exit 69
