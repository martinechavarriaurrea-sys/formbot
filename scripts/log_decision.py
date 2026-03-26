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
    parser = argparse.ArgumentParser(description="Registrar decision tecnica en docs/decisions.")
    parser.add_argument("--title", required=True, help="Titulo de la decision")
    parser.add_argument("--context", required=True, help="Contexto del problema")
    parser.add_argument(
        "--alternatives",
        nargs="+",
        required=True,
        help="Alternativas consideradas (separadas por espacios)",
    )
    parser.add_argument("--decision", required=True, help="Opcion elegida")
    parser.add_argument("--justification", required=True, help="Justificacion")
    parser.add_argument("--impact", required=True, help="Impacto en modulos")
    parser.add_argument("--status", default="Aprobada", help="Estado de la decision")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    traceability = TraceabilityRegistry(PROJECT_ROOT)
    traceability.append_decision(
        title=args.title,
        context=args.context,
        alternatives=args.alternatives,
        decision=args.decision,
        justification=args.justification,
        impact=args.impact,
        status=args.status,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

