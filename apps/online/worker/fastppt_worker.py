#!/usr/bin/env python3
"""
FastPPT Online - Contract Worker

Validate project-scoped worker tasks without granting arbitrary path access.

Usage:
    python fastppt_worker.py

Examples:
    echo '{"kind":"health"}' | python fastppt_worker.py

Dependencies:
    None (only uses standard library)
"""

from __future__ import annotations

import hashlib
import json
import sys
from typing import Any


ALLOWED_TASKS = {"health", "contract.validate", "prompt.snapshot", "qa.check"}


def _response(payload: dict[str, Any]) -> None:
    print(json.dumps(payload, ensure_ascii=False, separators=(",", ":")))


def main() -> int:
    sys.stdin.reconfigure(encoding="utf-8", errors="strict")
    sys.stdout.reconfigure(encoding="utf-8", errors="strict")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    try:
        task = json.load(sys.stdin)
    except (json.JSONDecodeError, UnicodeError) as exc:
        print(f"Invalid UTF-8 JSON task: {exc}", file=sys.stderr)
        return 2
    kind = str(task.get("kind", ""))
    if kind not in ALLOWED_TASKS:
        print(f"Unsupported worker task: {kind}", file=sys.stderr)
        return 2
    if kind == "health":
        _response({"ok": True, "worker": "fastppt-contract", "version": "0.1.0"})
        return 0
    payload = task.get("payload") or {}
    serialized = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    _response({"ok": True, "kind": kind, "snapshot_hash": hashlib.sha256(serialized.encode("utf-8")).hexdigest()})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
