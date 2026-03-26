from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from formbot.shared.utils import TraceabilityRegistry


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Registrar cambio manual en docs/changes.")
    parser.add_argument("--change", required=True, help="Descripcion del cambio")
    parser.add_argument("--reason", required=True, help="Motivo")
    parser.add_argument(
        "--files",
        nargs="+",
        required=True,
        help="Lista de archivos afectados",
    )
    parser.add_argument("--risks", required=True, help="Riesgos asociados")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    traceability = TraceabilityRegistry(PROJECT_ROOT)
    traceability.append_change(
        change=args.change,
        reason=args.reason,
        affected_files=args.files,
        risks=args.risks,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

