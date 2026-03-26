from __future__ import annotations

import hashlib
import json
import logging
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Iterable, Sequence

LOGGER = logging.getLogger(__name__)


def configure_logging(level: str = "INFO") -> None:
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    )


def normalize_text(value: str) -> str:
    lowered = value.strip().lower()
    without_accents = "".join(
        char
        for char in unicodedata.normalize("NFKD", lowered)
        if not unicodedata.combining(char)
    )
    normalized_separators = re.sub(r"[^a-z0-9]+", " ", without_accents)
    return " ".join(normalized_separators.split())


class TraceabilityRegistry:
    def __init__(self, project_root: Path) -> None:
        self._project_root = project_root
        self._docs_root = project_root / "docs"
        self._decisions_dir = self._docs_root / "decisions"
        self._changes_dir = self._docs_root / "changes"
        self._runs_dir = self._docs_root / "runs"

        self.decision_log_path = self._decisions_dir / "decision_log.md"
        self.change_log_path = self._changes_dir / "change_log.md"
        self.execution_log_path = self._runs_dir / "execution_log.md"
        self.architecture_doc_path = self._docs_root / "architecture.md"
        self._script_state_path = self._changes_dir / ".script_hash_state.json"

    def ensure_structure(self) -> None:
        self._decisions_dir.mkdir(parents=True, exist_ok=True)
        self._changes_dir.mkdir(parents=True, exist_ok=True)
        self._runs_dir.mkdir(parents=True, exist_ok=True)

        self._ensure_markdown(self.decision_log_path, "# Registro de Decisiones\n")
        self._ensure_markdown(self.change_log_path, "# Registro de Cambios\n")
        self._ensure_markdown(self.execution_log_path, "# Registro de Ejecuciones\n")
        self._ensure_markdown(self.architecture_doc_path, "# Arquitectura\n")

    def append_decision(
        self,
        title: str,
        context: str,
        alternatives: Sequence[str],
        decision: str,
        justification: str,
        impact: str,
        status: str,
    ) -> None:
        self.ensure_structure()
        numbered_alternatives = "\n".join(
            f"{index}. {alternative}"
            for index, alternative in enumerate(alternatives, start=1)
        )
        entry = (
            f"## [{self._timestamp()}] - Decision: {title}\n\n"
            f"### Contexto\n{context}\n\n"
            f"### Alternativas\n{numbered_alternatives}\n\n"
            f"### Decision\n{decision}\n\n"
            f"### Justificacion\n{justification}\n\n"
            f"### Impacto\n{impact}\n\n"
            f"### Estado\n{status}\n\n"
        )
        self._append_markdown(self.decision_log_path, entry)

    def append_change(
        self,
        change: str,
        reason: str,
        affected_files: Sequence[str],
        risks: str,
    ) -> None:
        self.ensure_structure()
        files_block = "\n".join(f"- {file_path}" for file_path in affected_files)
        entry = (
            f"## [{self._timestamp()}]\n\n"
            f"### Cambio\n{change}\n\n"
            f"### Motivo\n{reason}\n\n"
            f"### Archivos afectados\n{files_block}\n\n"
            f"### Riesgos\n{risks}\n\n"
        )
        self._append_markdown(self.change_log_path, entry)

    def append_execution(self, script_name: str, result: str, observations: str) -> None:
        self.ensure_structure()
        entry = (
            f"## [{self._timestamp()}]\n\n"
            f"### Script ejecutado\n{script_name}\n\n"
            f"### Resultado\n{result}\n\n"
            f"### Observaciones\n{observations}\n\n"
        )
        self._append_markdown(self.execution_log_path, entry)

    def register_script_changes(self, script_paths: Sequence[Path]) -> list[Path]:
        self.ensure_structure()
        state = self._load_script_state()
        changed_scripts: list[Path] = []

        for script_path in script_paths:
            absolute_path = (
                script_path
                if script_path.is_absolute()
                else (self._project_root / script_path).resolve()
            )
            if not absolute_path.exists():
                LOGGER.warning(
                    "No se pudo registrar cambio automatico: script no encontrado %s",
                    absolute_path,
                )
                continue

            digest = hashlib.sha256(absolute_path.read_bytes()).hexdigest()
            relative_path = self._to_project_relative(absolute_path)
            previous_digest = state.get(relative_path)

            if previous_digest == digest:
                continue

            change_message = (
                f"Registro inicial de script principal {relative_path}"
                if previous_digest is None
                else f"Cambio detectado automaticamente en script principal {relative_path}"
            )
            reason = (
                "Alta inicial en el monitor automatico de scripts principales"
                if previous_digest is None
                else "El hash del script cambio respecto a la ultima ejecucion registrada"
            )
            self.append_change(
                change=change_message,
                reason=reason,
                affected_files=[relative_path],
                risks="Cambios en scripts principales pueden alterar el comportamiento del pipeline.",
            )

            state[relative_path] = digest
            changed_scripts.append(absolute_path)

        self._save_script_state(state)
        return changed_scripts

    @staticmethod
    def _append_markdown(file_path: Path, entry: str) -> None:
        with file_path.open("a", encoding="utf-8") as fp:
            fp.write(entry.rstrip() + "\n\n")

    @staticmethod
    def _ensure_markdown(file_path: Path, header: str) -> None:
        if file_path.exists():
            return
        with file_path.open("w", encoding="utf-8") as fp:
            fp.write(header)

    @staticmethod
    def _timestamp() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _load_script_state(self) -> dict[str, str]:
        if not self._script_state_path.exists():
            return {}
        try:
            return json.loads(self._script_state_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            LOGGER.warning(
                "No se pudo leer estado de hash de scripts en %s. Se reinicia estado.",
                self._script_state_path,
            )
            return {}

    def _save_script_state(self, state: dict[str, str]) -> None:
        self._script_state_path.parent.mkdir(parents=True, exist_ok=True)
        self._script_state_path.write_text(
            json.dumps(state, indent=2, sort_keys=True),
            encoding="utf-8",
        )

    def _to_project_relative(self, absolute_path: Path) -> str:
        try:
            return (
                absolute_path.resolve()
                .relative_to(self._project_root.resolve())
                .as_posix()
            )
        except ValueError:
            return str(absolute_path.resolve())


def format_trace_lines(lines: Iterable[str]) -> str:
    cleaned = [line.strip() for line in lines if line.strip()]
    return "\n".join(cleaned) if cleaned else "Sin observaciones"


def find_duplicates(values: Sequence[str]) -> set[str]:
    """Retorna el conjunto de valores que aparecen mas de una vez en la secuencia."""
    seen: set[str] = set()
    duplicated: set[str] = set()
    for value in values:
        if value in seen:
            duplicated.add(value)
            continue
        seen.add(value)
    return duplicated
