#!/usr/bin/env python3
from __future__ import annotations

import re
from pathlib import Path

from build_book import apply_overrides, extract_docx


ROOT = Path(__file__).resolve().parent
BUILD_BOOK = ROOT / "build_book.py"


def collect_used_indices() -> set[int]:
    text = BUILD_BOOK.read_text(encoding="utf-8")
    used: set[int] = set()

    for match in re.finditer(r"p\[(\d+)\]", text):
        used.add(int(match.group(1)))

    for line in text.splitlines():
        line = line.strip()
        match = re.match(r"for idx in range\((\d+),\s*(\d+)\):", line)
        if match:
            start, end = map(int, match.groups())
            used.update(range(start, end))
            continue

        match = re.match(r"for idx in \(([^)]+)\):", line)
        if match:
            nums = [int(x.strip()) for x in match.group(1).split(",") if x.strip()]
            used.update(nums)
            used.update(i + 1 for i in nums)

    return used


def main() -> None:
    paras, _ = extract_docx()
    paras = apply_overrides(paras)
    used = collect_used_indices()

    missing = [i for i in sorted(paras) if i not in used]
    print(f"paragraphs_in_docx={len(paras)}")
    print(f"paragraphs_referenced={len(used)}")

    if not missing:
        print("audit=ok")
        return

    print("audit=missing")
    for idx in missing:
        print(f"{idx:03d}: {paras[idx]}")


if __name__ == "__main__":
    main()
