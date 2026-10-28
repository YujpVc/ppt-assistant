#!/usr/bin/env bash
set -euo pipefail

# Run the installed pptfast package without importing or executing workspace code.
# Desktop deployments historically keep it in ~/ppt-assistant/node_modules; the
# repository checkout keeps the same package layout beside this launcher.
engine_root=$(CDPATH= cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
candidate_paths=(
  "$engine_root/node_modules/@liustack/pptfast/dist/cli.js"
  "${HOME}/ppt-assistant/node_modules/@liustack/pptfast/dist/cli.js"
)

for cli in "${candidate_paths[@]}"; do
  if [[ -f "$cli" ]]; then
    if ! command -v node >/dev/null 2>&1; then
      echo "pptfast: 找到运行包但未找到 Node.js（需要 Node 22.19+）" >&2
      exit 78
    fi
    exec node "$cli" "$@"
  fi
done

cat >&2 <<'EOF'
pptfast: 未找到已安装的 @liustack/pptfast 运行包。
请在部署环境安装 @liustack/pptfast@0.20.0，或使用内置 build_deck.py；不要执行用户工作区脚本。
EOF
exit 78
