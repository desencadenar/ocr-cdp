"""Lectura completa guardada de cada documento, para volver a detectar sin repetir el OCR.

Con `guardar_lectura_completa` (por defecto true) la corrida guarda en `.ocr_fiscal/lecturas.sqlite3`, por cada
documento leído:
- el texto y las líneas de cada página, con coordenadas, confianza y origen (texto del PDF u OCR);
- los QR leídos;
- las relecturas puntuales de línea que hizo Tesseract.

La clave es el hash SHA-256 del contenido del documento junto con los ajustes que cambian la lectura (idiomas,
resolución, máximo de píxeles), así que un documento modificado o leído con otros ajustes no se confunde con otro.

El comando `redetectar` usa estas lecturas para repetir la extracción y el Excel sin Tesseract.

Aviso de privacidad: es la transcripción completa de cada documento. Quien reciba la carpeta de salida la recibe.
"""

from __future__ import annotations

import contextlib
from dataclasses import asdict
import gzip
import hashlib
import json
from pathlib import Path
import sqlite3
import threading
from typing import Any

from .config import AppConfig
from .documents import DocumentRead, PageRead
from .models import LayoutLine, StageEvent

CACHE_FILENAME = "lecturas.sqlite3"
_SCHEMA_VERSION = "1"


def reading_signature(config: AppConfig) -> str:
    """Ajustes que cambian lo que se lee de un documento."""
    return f"v{_SCHEMA_VERSION}|{config.languages}|{config.render_dpi}|{config.max_pixels}"


def unit_key(content: bytes, config: AppConfig) -> str:
    return hashlib.sha256(content).hexdigest() + "|" + reading_signature(config)


def _page_to_dict(page: PageRead) -> dict[str, Any]:
    data = asdict(page)
    data["layout_lines"] = [line.as_dict() for line in page.layout_lines]
    return data


def _page_from_dict(data: dict[str, Any]) -> PageRead:
    lines = [
        LayoutLine(
            line_number=item["line_number"], text=item["text"], block_number=item.get("block_number", 1),
            bbox=tuple(item["bbox"]) if item.get("bbox") is not None else None,
            ocr_confidence=item.get("ocr_confidence"), source=item.get("source", "native_pdf_text"),
        )
        for item in data.get("layout_lines", [])
    ]
    fields = {k: v for k, v in data.items() if k != "layout_lines"}
    return PageRead(**fields, layout_lines=lines)


def serialize(read: DocumentRead, rereads: dict[str, str | None]) -> bytes:
    payload = {
        "path": str(read.path),
        "source_format": read.source_format,
        "page_count": read.page_count,
        "pages": [_page_to_dict(p) for p in read.pages],
        "events": [e.as_dict() for e in read.events],
        "status": read.status,
        "ocr_engine": read.ocr_engine,
        "ocr_version": read.ocr_version,
        "rereads": rereads,
    }
    return gzip.compress(json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8"))


def _event_from_dict(data: dict[str, Any]) -> StageEvent:
    known = {"stage", "status", "message", "page_number", "duration_ms", "source_id", "member_path", "details"}
    kwargs = {k: v for k, v in data.items() if k in known}
    if data.get("container_chain") is not None:
        kwargs["container_chain"] = tuple(data["container_chain"])
    return StageEvent(**kwargs)


def deserialize(blob: bytes) -> tuple[DocumentRead, dict[str, str | None]]:
    data = json.loads(gzip.decompress(blob).decode("utf-8"))
    read = DocumentRead(
        path=Path(data["path"]),
        source_format=data["source_format"],
        page_count=data["page_count"],
        pages=[_page_from_dict(p) for p in data["pages"]],
        events=[_event_from_dict(e) for e in data.get("events", [])],
        status=data.get("status", "read"),
        ocr_engine=data.get("ocr_engine", "none"),
        ocr_version=data.get("ocr_version", "none"),
    )
    return read, dict(data.get("rereads") or {})


def reread_key(page_number: int, bbox: Any) -> str:
    return f"{page_number}|{','.join(f'{float(v):.2f}' for v in bbox)}"


class ReadingCache:
    """Base SQLite con una fila por documento leído. Abre y cierra la conexión en cada operación (así no deja el
    archivo bloqueado en Windows) y se puede usar desde varios hilos."""

    def __init__(self, path: str | Path, *, readonly: bool = False) -> None:
        self.path = Path(path)
        self.readonly = readonly
        self._lock = threading.Lock()
        if readonly:
            if not self.path.is_file():
                raise ValueError(f"No hay lecturas guardadas en {self.path}.")
        else:
            with self._connect() as conn:
                conn.execute("CREATE TABLE IF NOT EXISTS lecturas (clave TEXT PRIMARY KEY, contenido BLOB NOT NULL)")

    @contextlib.contextmanager
    def _connect(self):
        if self.readonly:
            conn = sqlite3.connect(f"file:{self.path.resolve().as_posix()}?mode=ro", uri=True)
        else:
            conn = sqlite3.connect(self.path)
        try:
            with conn:
                yield conn
        finally:
            conn.close()

    def put(self, key: str, read: DocumentRead, rereads: dict[str, str | None]) -> None:
        blob = serialize(read, rereads)
        with self._lock, self._connect() as conn:
            conn.execute("INSERT OR REPLACE INTO lecturas (clave, contenido) VALUES (?, ?)", (key, blob))

    def get(self, key: str) -> tuple[DocumentRead, dict[str, str | None]] | None:
        with self._lock, self._connect() as conn:
            row = conn.execute("SELECT contenido FROM lecturas WHERE clave = ?", (key,)).fetchone()
        return deserialize(row[0]) if row else None

    def count(self) -> int:
        with self._lock, self._connect() as conn:
            return int(conn.execute("SELECT COUNT(*) FROM lecturas").fetchone()[0])

    def first(self) -> DocumentRead:
        with self._lock, self._connect() as conn:
            row = conn.execute("SELECT contenido FROM lecturas LIMIT 1").fetchone()
        return deserialize(row[0])[0]

    def close(self) -> None:
        """Nada que cerrar: cada operación abre y cierra su conexión. Se conserva por compatibilidad."""
