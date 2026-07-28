#!/usr/bin/env python3
"""Standalone CLI entry point for exporting the inventory dashboard to PPTX."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import shutil
import subprocess
import sys


def find_node(explicit: str | None) -> str:
    candidates = [
        explicit,
        os.environ.get("NODE_BINARY"),
        shutil.which("node"),
        str(Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node"),
        str(Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/bin/node"),
        str(Path.home() / ".cache/codex-runtimes/codex-primary-runtime/node/bin/node"),
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return candidate
    raise SystemExit("找不到 Node.js。請使用 --node 指定 Node 22 以上的執行檔。")


def main() -> int:
    parser = argparse.ArgumentParser(description="匯出 My Inventory Intelligence 管理層儀表板 PPTX")
    parser.add_argument("--input", required=True, help="Dashboard JSON 檔案")
    parser.add_argument("--output", required=True, help="輸出的 PPTX 檔案")
    parser.add_argument("--node", help="Node.js 執行檔路徑（選填）")
    args = parser.parse_args()

    root = Path(__file__).resolve().parents[1]
    command = [
        find_node(args.node),
        str(root / "scripts" / "export_inventory_pptx.mjs"),
        "--input",
        str(Path(args.input).resolve()),
        "--output",
        str(Path(args.output).resolve()),
    ]
    completed = subprocess.run(command, cwd=root, check=False)
    return completed.returncode


if __name__ == "__main__":
    sys.exit(main())
