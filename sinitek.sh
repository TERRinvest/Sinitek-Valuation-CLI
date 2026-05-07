#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)"
script_path="$script_dir/sinitek.ps1"

if command -v cygpath >/dev/null 2>&1; then
  script_path="$(cygpath -w "$script_path")"
elif command -v wslpath >/dev/null 2>&1; then
  script_path="$(wslpath -w "$script_path")"
fi

if ! command -v powershell.exe >/dev/null 2>&1; then
  echo "ERROR: powershell.exe is required. Run this on Windows, Git Bash, or WSL with Windows interop enabled." >&2
  exit 127
fi

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$script_path" "$@"
