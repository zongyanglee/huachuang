from __future__ import annotations

import ast
from pathlib import Path
import sys
import tokenize


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCAN_ROOTS = (PROJECT_ROOT / "src", PROJECT_ROOT / "scripts")


def main() -> int:
    failures: list[tuple[Path, Exception]] = []
    checked = 0
    for root in SCAN_ROOTS:
        for path in sorted(root.rglob("*.py")):
            if "__pycache__" in path.parts:
                continue
            checked += 1
            try:
                with tokenize.open(path) as handle:
                    ast.parse(handle.read(), filename=str(path))
            except (OSError, SyntaxError, UnicodeError) as exc:
                failures.append((path, exc))

    print(f"Checked Python files: {checked}")
    if not failures:
        print("Syntax errors: 0")
        return 0

    print(f"Syntax errors: {len(failures)}")
    for path, exc in failures:
        print(f"- {path.relative_to(PROJECT_ROOT)}: {exc}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
