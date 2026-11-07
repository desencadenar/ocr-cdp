from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone
import json
from pathlib import Path
import sqlite3
from typing import Any, Iterator, Mapping, Sequence

from .models import (
    ArchiveSummary,
    IDENTITY_COMPLETE,
    IDENTITY_INCOMPLETE,
    IDENTITY_NOT_APPLICABLE,
    SourceRef,
    StageEvent,
    normalize_detection_status,
)


# Columnas de la base que guardan listas o diccionarios como texto JSON.
# Cada una dice con qué nombre se devuelve al leerla y cuál es su valor vacío.
_JSON_COLUMNS: dict[str, tuple[str, str]] = {
    "container_chain_json": ("container_chain", "[]"),
    "fields_original_json": ("fields_original", "{}"),
    "fields_normalized_json": ("fields_normalized", "{}"),
    "field_evidence_json": ("field_evidence", "{}"),
    "field_statuses_json": ("field_statuses", "{}"),
    "confidence_json": ("confidence", "{}"),
    "candidates_json": ("candidates", "{}"),
    "observations_json": ("observations", "[]"),
    "read_error_details_json": ("read_error_details", "[]"),
    "line_items_json": ("line_items", "[]"),
    "derived_checks_json": ("derived_checks", "[]"),
    "third_party_rucs_json": ("third_party_rucs", "[]"),
    "secondary_review_fields_json": ("secondary_review_fields", "[]"),
}

# Columnas mínimas que tiene que tener la tabla documents. Si una base antigua no las trae, se rechaza
# para no mezclar versiones distintas.
_REQUIRED_COLUMNS = {
    "id",
    "original_path",
    "source_format",
    "page_start",
    "page_end",
    "detection_status",
    "field_evidence_json",
    "read_error",
}

# Columnas que se fueron añadiendo con el tiempo. Si una base vieja no las tiene, se agregan al abrirla
# (con su valor por omisión) para poder seguir leyéndola.
_DOCUMENTS_EXTRA_COLUMNS: tuple[tuple[str, str], ...] = (
    ("identity_status", "TEXT NOT NULL DEFAULT 'no_aplica'"),
    ("detection_reason", "TEXT NOT NULL DEFAULT ''"),
    ("detection_excerpt", "TEXT"),
    ("identity_reason", "TEXT NOT NULL DEFAULT ''"),
    ("identity_excerpt", "TEXT"),
    ("third_party_rucs_json", "TEXT NOT NULL DEFAULT '[]'"),
    ("candidate_fiscal_key", "TEXT"),
    ("secondary_review_fields_json", "TEXT NOT NULL DEFAULT '[]'"),
    # Datos de la verificación oficial. Este programa no consulta a SUNAT, así que el estado queda en «pendiente».
    ("official_validation_status", "TEXT NOT NULL DEFAULT 'pendiente'"),
    ("official_validation_date", "TEXT"),
    ("official_validation_source", "TEXT"),
    ("official_validation_reference", "TEXT"),
    ("official_validation_reason", "TEXT NOT NULL DEFAULT ''"),
    # Canal que capturó la respuesta oficial. Una base creada por una versión anterior se
    # migra con el valor por defecto, de modo que su estado persistido no publica.
    ("official_validation_channel", "TEXT NOT NULL DEFAULT 'importado_sin_verificar'"),
    # Resultado de contrastar el emisor y el receptor con el padrón RUC local (si se configuró uno).
    ("emitter_padron_status", "TEXT NOT NULL DEFAULT 'no_aplica'"),
    ("emitter_padron_razon_social", "TEXT"),
    ("emitter_padron_estado", "TEXT"),
    ("emitter_padron_cond_domicilio", "TEXT"),
    ("emitter_padron_cutoff_date", "TEXT"),
    ("emitter_padron_source", "TEXT"),
    ("emitter_name_discrepancy", "INTEGER NOT NULL DEFAULT 0"),
    ("recipient_padron_status", "TEXT NOT NULL DEFAULT 'no_aplica'"),
    ("recipient_padron_razon_social", "TEXT"),
    ("recipient_padron_estado", "TEXT"),
    ("recipient_padron_cond_domicilio", "TEXT"),
    # Resultado de reunir las copias del mismo comprobante (misma clave fiscal) y cuál de ellas es la principal.
    ("key_reconciliation_status", "TEXT NOT NULL DEFAULT 'sin_clave_completa'"),
    ("key_reconciliation_reason", "TEXT NOT NULL DEFAULT ''"),
    ("is_primary_representation", "INTEGER NOT NULL DEFAULT 0"),
    ("eligible_for_valid_sheet", "INTEGER NOT NULL DEFAULT 0"),
)

# Lo mismo para la tabla de los ZIP.
_ARCHIVE_EXTRA_COLUMNS: tuple[tuple[str, str], ...] = (
    ("inspection_status", "TEXT NOT NULL DEFAULT 'completo'"),
    ("processing_status", "TEXT NOT NULL DEFAULT 'pendiente'"),
    ("processed_members", "INTEGER NOT NULL DEFAULT 0"),
    ("unprocessed_members_count", "INTEGER NOT NULL DEFAULT 0"),
    ("excluded_xml_members", "INTEGER NOT NULL DEFAULT 0"),
    ("cancelled", "INTEGER NOT NULL DEFAULT 0"),
)


# Fecha y hora actuales en UTC, en formato ISO.
def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


# La base de datos de la corrida: un archivo SQLite dentro de la carpeta de salida (.ocr_fiscal/state.sqlite3).
# Todo lo que se lee de los documentos queda aquí, y el Excel se genera después a partir de ella.
class StateStore:
    """Guarda en SQLite los resultados de esta corrida (comprobantes, inventario de archivos, detalle y eventos).

    No se usa como caché ni como fuente de inferencia.
    """

    # Crea la carpeta si hace falta y prepara las tablas.
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    # Abre una conexión por cada operación y la cierra siempre al terminar.
    @contextmanager
    def _connection(self) -> Iterator[sqlite3.Connection]:
        connection = sqlite3.connect(self.path, timeout=30)
        try:
            connection.row_factory = sqlite3.Row
            connection.execute("PRAGMA foreign_keys=ON")
            with connection:
                yield connection
        finally:
            # La transacción termina al salir de "with connection", pero en Windows
            # el archivo sigue ocupado hasta cerrar también la conexión.
            connection.close()

    # Crea las tablas si no existen y completa las que les falten columnas.
    # Si encuentra una base de un esquema demasiado antiguo, se detiene en lugar de mezclarla.
    def _initialize(self) -> None:
        with self._connection() as connection:
            # Paso 1: si ya hay una tabla documents, comprobamos que es de un esquema que sabemos leer.
            tables = {
                row[0]
                for row in connection.execute("SELECT name FROM sqlite_master WHERE type='table'")
            }
            if "documents" in tables:
                existing = {row[1] for row in connection.execute("PRAGMA table_info(documents)")}
                if not _REQUIRED_COLUMNS.issubset(existing):
                    raise ValueError(
                        "La base pertenece a un esquema anterior; elija una carpeta de salida nueva para conservarla intacta"
                    )
            # Paso 2: crear las tablas (metadatos de la corrida, inventario de archivos, ZIP, documentos,
            # detalle de ítems y eventos) y sus índices.
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS run_metadata (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    run_id TEXT NOT NULL,
                    schema_version INTEGER NOT NULL DEFAULT 1,
                    target_recipient_ruc TEXT,
                    run_config_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS source_inventory (
                    source_id INTEGER PRIMARY KEY,
                    source_path TEXT NOT NULL,
                    container_chain_json TEXT NOT NULL DEFAULT '[]',
                    member_path TEXT,
                    nested_depth INTEGER NOT NULL DEFAULT 0,
                    detected_format TEXT NOT NULL,
                    inventory_status TEXT NOT NULL,
                    status_reason TEXT NOT NULL DEFAULT '',
                    action_taken TEXT NOT NULL DEFAULT '',
                    byte_size INTEGER NOT NULL DEFAULT 0,
                    compressed_size INTEGER,
                    display_location TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_inventory_status
                    ON source_inventory(inventory_status);

                CREATE TABLE IF NOT EXISTS archive_containers (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    source_id INTEGER NOT NULL,
                    source_path TEXT NOT NULL,
                    container_chain_json TEXT NOT NULL DEFAULT '[]',
                    nested_depth INTEGER NOT NULL DEFAULT 0,
                    status TEXT NOT NULL,
                    inspection_status TEXT NOT NULL DEFAULT 'completo',
                    processing_status TEXT NOT NULL DEFAULT 'pendiente',
                    direct_members INTEGER NOT NULL DEFAULT 0,
                    recursive_members INTEGER NOT NULL DEFAULT 0,
                    directory_entries INTEGER NOT NULL DEFAULT 0,
                    supported_members INTEGER NOT NULL DEFAULT 0,
                    processed_members INTEGER NOT NULL DEFAULT 0,
                    unprocessed_members_count INTEGER NOT NULL DEFAULT 0,
                    excluded_xml_members INTEGER NOT NULL DEFAULT 0,
                    unsupported_members INTEGER NOT NULL DEFAULT 0,
                    encrypted_members INTEGER NOT NULL DEFAULT 0,
                    traversal_blocked_members INTEGER NOT NULL DEFAULT 0,
                    corrupted INTEGER NOT NULL DEFAULT 0,
                    error_reason TEXT,
                    limit_reached INTEGER NOT NULL DEFAULT 0,
                    limit_reasons_json TEXT NOT NULL DEFAULT '[]',
                    pending_members_json TEXT NOT NULL DEFAULT '[]',
                    cancelled INTEGER NOT NULL DEFAULT 0,
                    max_nested_depth_observed INTEGER NOT NULL DEFAULT 1,
                    detected_receipts INTEGER NOT NULL DEFAULT 0,
                    ambiguous_records INTEGER NOT NULL DEFAULT 0,
                    non_receipts INTEGER NOT NULL DEFAULT 0,
                    read_errors INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS documents (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    source_id INTEGER,
                    original_path TEXT NOT NULL,
                    container_chain_json TEXT NOT NULL DEFAULT '[]',
                    member_path TEXT,
                    nested_depth INTEGER NOT NULL DEFAULT 0,
                    display_location TEXT NOT NULL DEFAULT '',
                    source_format TEXT NOT NULL,
                    reading_source TEXT NOT NULL DEFAULT 'native_pdf_text',
                    page_count INTEGER NOT NULL DEFAULT 0,
                    page_start INTEGER NOT NULL DEFAULT 1,
                    page_end INTEGER NOT NULL DEFAULT 1,
                    block_start INTEGER NOT NULL DEFAULT 1,
                    block_end INTEGER NOT NULL DEFAULT 1,
                    detection_status TEXT NOT NULL,
                    identity_status TEXT NOT NULL DEFAULT 'no_aplica',
                    comprobante_detectado INTEGER NOT NULL DEFAULT 0,
                    identidad_completa INTEGER NOT NULL DEFAULT 0,
                    renombrable INTEGER NOT NULL DEFAULT 0,
                    classification_rule TEXT NOT NULL DEFAULT '',
                    non_receipt_type TEXT,
                    decision_comment TEXT NOT NULL DEFAULT '',
                    detection_reason TEXT NOT NULL DEFAULT '',
                    detection_excerpt TEXT,
                    identity_reason TEXT NOT NULL DEFAULT '',
                    identity_excerpt TEXT,
                    recipient_ruc_match INTEGER NOT NULL DEFAULT 0,
                    recipient_role_status TEXT NOT NULL DEFAULT 'not_found',
                    issuer_role_status TEXT NOT NULL DEFAULT 'not_found',
                    third_party_rucs_json TEXT NOT NULL DEFAULT '[]',
                    candidate_fiscal_key TEXT,
                    secondary_review_fields_json TEXT NOT NULL DEFAULT '[]',
                    general_description TEXT,
                    requires_human_review INTEGER NOT NULL DEFAULT 1,
                    line_items_json TEXT NOT NULL DEFAULT '[]',
                    derived_checks_json TEXT NOT NULL DEFAULT '[]',
                    field_statuses_json TEXT NOT NULL DEFAULT '{}',
                    fields_original_json TEXT NOT NULL DEFAULT '{}',
                    fields_normalized_json TEXT NOT NULL DEFAULT '{}',
                    field_evidence_json TEXT NOT NULL DEFAULT '{}',
                    confidence_json TEXT NOT NULL DEFAULT '{}',
                    candidates_json TEXT NOT NULL DEFAULT '{}',
                    observations_json TEXT NOT NULL DEFAULT '[]',
                    read_error INTEGER NOT NULL DEFAULT 0,
                    read_error_details_json TEXT NOT NULL DEFAULT '[]',
                    extracted_text TEXT NOT NULL DEFAULT '',
                    text_sidecar_path TEXT,
                    normalized_filename TEXT,
                    organized_path TEXT,
                    created_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_documents_detection_status
                    ON documents(detection_status);
                CREATE INDEX IF NOT EXISTS idx_documents_original_path
                    ON documents(original_path);

                CREATE TABLE IF NOT EXISTS document_line_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    document_id INTEGER NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
                    item_index INTEGER NOT NULL,
                    code TEXT,
                    description_observed TEXT,
                    description_normalized TEXT,
                    quantity_observed TEXT,
                    quantity_normalized TEXT,
                    unit_observed TEXT,
                    unit_normalized TEXT,
                    unit_price_observed TEXT,
                    unit_price_normalized TEXT,
                    line_total_observed TEXT,
                    line_total_normalized TEXT,
                    page_number INTEGER NOT NULL DEFAULT 1,
                    block_number INTEGER NOT NULL DEFAULT 1,
                    line_number INTEGER,
                    bbox_json TEXT,
                    excerpt TEXT NOT NULL DEFAULT '',
                    source TEXT NOT NULL DEFAULT 'native_pdf_text',
                    ocr_confidence REAL,
                    heuristic_confidence REAL,
                    missing_reason TEXT
                );
                CREATE INDEX IF NOT EXISTS idx_line_items_document_id
                    ON document_line_items(document_id, item_index);

                CREATE TABLE IF NOT EXISTS processing_events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    document_id INTEGER REFERENCES documents(id) ON DELETE CASCADE,
                    source_id INTEGER,
                    member_path TEXT,
                    container_chain_json TEXT NOT NULL DEFAULT '[]',
                    stage TEXT NOT NULL,
                    status TEXT NOT NULL,
                    message TEXT NOT NULL DEFAULT '',
                    page_number INTEGER,
                    duration_ms REAL,
                    details_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_events_document_id
                    ON processing_events(document_id, id);
                PRAGMA user_version=4;
                """
            )
            # Paso 3: migrar bases de versiones anteriores añadiéndoles las columnas nuevas que no tengan.
            existing_doc_cols = {row[1] for row in connection.execute("PRAGMA table_info(documents)")}
            for col_name, col_def in _DOCUMENTS_EXTRA_COLUMNS:
                if col_name not in existing_doc_cols:
                    connection.execute(f"ALTER TABLE documents ADD COLUMN {col_name} {col_def}")
            existing_arch_cols = {row[1] for row in connection.execute("PRAGMA table_info(archive_containers)")}
            for col_name, col_def in _ARCHIVE_EXTRA_COLUMNS:
                if col_name not in existing_arch_cols:
                    connection.execute(f"ALTER TABLE archive_containers ADD COLUMN {col_name} {col_def}")

    # Convierte una fila de la base en un diccionario listo para usar: vuelve a abrir los JSON,
    # pasa los 0/1 a verdadero/falso y rellena con valores por omisión lo que las bases antiguas no tenían.
    @staticmethod
    def _decode_row(row: sqlite3.Row | Mapping[str, Any]) -> dict[str, Any]:
        # Paso 1: los campos guardados como JSON vuelven a ser listas y diccionarios.
        item = dict(row)
        for column, (key, default_json) in _JSON_COLUMNS.items():
            if column in item:
                raw = item.pop(column)
                item[key] = json.loads(raw if raw else default_json)
        # Paso 2: normalizamos los estados y pasamos los 0/1 de SQLite a verdadero/falso.
        item["detection_status"] = normalize_detection_status(item.get("detection_status"))
        item["comprobante_detectado"] = bool(item.get("comprobante_detectado"))
        item["identidad_completa"] = bool(item.get("identidad_completa"))
        # Si una base antigua no traía el estado de identidad, se deduce de si estaba completa.
        if not item.get("identity_status") or (item.get("comprobante_detectado") and item.get("identity_status") == "no_aplica"):
            if item["comprobante_detectado"]:
                item["identity_status"] = IDENTITY_COMPLETE if item["identidad_completa"] else IDENTITY_INCOMPLETE
            else:
                item["identity_status"] = IDENTITY_NOT_APPLICABLE
        item["renombrable"] = bool(item.get("renombrable"))
        item["read_error"] = bool(item.get("read_error"))
        item["recipient_ruc_match"] = bool(item.get("recipient_ruc_match"))
        item["requires_human_review"] = bool(item.get("requires_human_review", True))
        item["emitter_name_discrepancy"] = bool(item.get("emitter_name_discrepancy", False))
        item["is_primary_representation"] = bool(item.get("is_primary_representation", False))
        item["eligible_for_valid_sheet"] = bool(item.get("eligible_for_valid_sheet", False))
        # Paso 3: valores por omisión para las columnas que una base antigua no tenía.
        item.setdefault("official_validation_status", "pendiente")
        item.setdefault("official_validation_date", None)
        item.setdefault("official_validation_source", None)
        item.setdefault("official_validation_reference", None)
        item.setdefault("official_validation_reason", "")
        item.setdefault("official_validation_channel", "importado_sin_verificar")
        item.setdefault("emitter_padron_status", "no_aplica")
        item.setdefault("emitter_padron_razon_social", None)
        item.setdefault("emitter_padron_estado", None)
        item.setdefault("emitter_padron_cond_domicilio", None)
        item.setdefault("emitter_padron_cutoff_date", None)
        item.setdefault("emitter_padron_source", None)
        item.setdefault("recipient_padron_status", "no_aplica")
        item.setdefault("recipient_padron_razon_social", None)
        item.setdefault("recipient_padron_estado", None)
        item.setdefault("recipient_padron_cond_domicilio", None)
        item.setdefault("key_reconciliation_status", "sin_clave_completa")
        item.setdefault("key_reconciliation_reason", "")
        # Rangos de página y de bloque, en forma de lista [inicio, fin].
        item["page_range"] = [item.get("page_start", 1), item.get("page_end", 1)]
        item["block_range"] = [item.get("block_start", 1), item.get("block_end", 1)]
        item.setdefault("third_party_rucs", [])
        item.setdefault("secondary_review_fields", [])
        return item

    # Guarda los datos de la corrida (RUC del comprador, ajustes, estado final). Devuelve el id del registro.
    def store_run_metadata(self, metadata: Mapping[str, Any]) -> int:
        now = _utc_now()
        with self._connection() as connection:
            cursor = connection.execute(
                """INSERT INTO run_metadata (
                    run_id, schema_version, target_recipient_ruc, run_config_json, created_at
                ) VALUES (?,?,?,?,?)""",
                (
                    str(metadata.get("run_id") or "default-run"),
                    int(metadata.get("schema_version") or 1),
                    metadata.get("target_recipient_ruc"),
                    json.dumps(dict(metadata), ensure_ascii=False),
                    now,
                ),
            )
            return int(cursor.lastrowid)

    # Devuelve los datos de la última corrida guardada, o None si no hay ninguna.
    def get_latest_run_metadata(self) -> dict[str, Any] | None:
        with self._connection() as connection:
            row = connection.execute("SELECT * FROM run_metadata ORDER BY id DESC LIMIT 1").fetchone()
        if row is None:
            return None
        data = dict(row)
        data["run_config"] = json.loads(data.pop("run_config_json") or "{}")
        return data

    # Guarda la lista de archivos encontrados (y de miembros de ZIP) con su formato y su estado.
    # Si un archivo ya estaba, se reemplaza.
    def store_inventory(self, entries: Sequence[SourceRef | Mapping[str, Any]]) -> None:
        if not entries:
            return
        now = _utc_now()
        rows = []
        for entry in entries:
            d = entry.as_dict() if isinstance(entry, SourceRef) else dict(entry)
            rows.append(
                (
                    int(d["source_id"]),
                    str(d["source_path"]),
                    json.dumps(list(d.get("container_chain") or []), ensure_ascii=False),
                    d.get("member_path"),
                    int(d.get("nested_depth") or 0),
                    str(d.get("detected_format") or "unknown"),
                    str(d.get("inventory_status") or "unknown"),
                    str(d.get("status_reason") or ""),
                    str(d.get("action_taken") or ""),
                    int(d.get("byte_size") or 0),
                    d.get("compressed_size"),
                    str(d.get("display_location") or d["source_path"]),
                    now,
                )
            )
        with self._connection() as connection:
            connection.executemany(
                """INSERT OR REPLACE INTO source_inventory (
                    source_id, source_path, container_chain_json, member_path,
                    nested_depth, detected_format, inventory_status, status_reason,
                    action_taken, byte_size, compressed_size, display_location, created_at
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                rows,
            )

    # Devuelve todo el inventario de archivos, en el orden en que se encontraron.
    def list_inventory(self) -> list[dict[str, Any]]:
        with self._connection() as connection:
            rows = connection.execute("SELECT * FROM source_inventory ORDER BY source_id").fetchall()
        result: list[dict[str, Any]] = []
        for row in rows:
            d = dict(row)
            d["container_chain"] = json.loads(d.pop("container_chain_json") or "[]")
            result.append(d)
        return result

    # Guarda el resumen de cada ZIP. Borra los anteriores: siempre quedan sólo los de la última pasada.
    def store_archive_summaries(self, summaries: Sequence[ArchiveSummary | Mapping[str, Any]]) -> None:
        if not summaries:
            return
        now = _utc_now()
        with self._connection() as connection:
            connection.execute("DELETE FROM archive_containers")
            for item in summaries:
                d = item.as_dict() if isinstance(item, ArchiveSummary) else dict(item)
                connection.execute(
                    """INSERT INTO archive_containers (
                        source_id, source_path, container_chain_json, nested_depth,
                        status, inspection_status, processing_status,
                        direct_members, recursive_members, directory_entries,
                        supported_members, processed_members, unprocessed_members_count,
                        excluded_xml_members, unsupported_members, encrypted_members,
                        traversal_blocked_members, corrupted, error_reason,
                        limit_reached, limit_reasons_json, pending_members_json,
                        cancelled, max_nested_depth_observed, detected_receipts,
                        ambiguous_records, non_receipts, read_errors, created_at
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        int(d.get("source_id") or 0),
                        str(d.get("source_path") or ""),
                        json.dumps(list(d.get("container_chain") or []), ensure_ascii=False),
                        int(d.get("nested_depth") or 0),
                        str(d.get("status") or "completo"),
                        str(d.get("inspection_status") or d.get("status") or "completo"),
                        str(d.get("processing_status") or "pendiente"),
                        int(d.get("direct_members") or 0),
                        int(d.get("recursive_members") or 0),
                        int(d.get("directory_entries") or 0),
                        int(d.get("supported_members") or 0),
                        int(d.get("processed_members") or 0),
                        int(d.get("unprocessed_members_count") or 0),
                        int(d.get("excluded_xml_members") or 0),
                        int(d.get("unsupported_members") or 0),
                        int(d.get("encrypted_members") or 0),
                        int(d.get("traversal_blocked_members") or 0),
                        int(bool(d.get("corrupted"))),
                        d.get("error_reason"),
                        int(bool(d.get("limit_reached"))),
                        json.dumps(list(d.get("limit_reasons") or []), ensure_ascii=False),
                        json.dumps(list(d.get("pending_members") or []), ensure_ascii=False),
                        int(bool(d.get("cancelled"))),
                        int(d.get("max_nested_depth_observed") or 1),
                        int(d.get("detected_receipts") or 0),
                        int(d.get("ambiguous_records") or 0),
                        int(d.get("non_receipts") or 0),
                        int(d.get("read_errors") or 0),
                        now,
                    ),
                )

    # DEPRECADO: nadie lo llama. El resumen de los ZIP sólo se escribe; no se vuelve a leer.
    def list_archive_summaries(self) -> list[dict[str, Any]]:
        with self._connection() as connection:
            rows = connection.execute("SELECT * FROM archive_containers ORDER BY id").fetchall()
        result: list[dict[str, Any]] = []
        for row in rows:
            d = dict(row)
            d["container_chain"] = json.loads(d.pop("container_chain_json") or "[]")
            d["limit_reasons"] = json.loads(d.pop("limit_reasons_json") or "[]")
            d["pending_members"] = json.loads(d.pop("pending_members_json") or "[]")
            d["corrupted"] = bool(d.get("corrupted"))
            d["limit_reached"] = bool(d.get("limit_reached"))
            d["cancelled"] = bool(d.get("cancelled"))
            result.append(d)
        return result

    # Guarda un documento detectado (un bloque de página) con todos sus datos y, si lo tiene,
    # el detalle de sus ítems. Devuelve el id que le asignó la base.
    def store_document(self, record: Mapping[str, Any]) -> int:
        now = _utc_now()
        # Paso 1: preparar los valores. Lo que falte se rellena con valores por omisión razonables
        # (por ejemplo, la página 1 si no viene indicada).
        container_chain = list(record.get("container_chain") or [])
        member_path = record.get("member_path")
        original_path = str(record["original_path"])
        display_location = str(
            record.get("display_location")
            or (f"{original_path} :: {' -> '.join(container_chain)}" if container_chain else original_path)
        )
        line_items = list(record.get("line_items") or [])
        # Los estados de detección y de identidad, si no vienen dados, se deducen de si es un comprobante
        # y de si su identidad está completa.
        det_status = normalize_detection_status(record.get("detection_status"))
        is_cdp = bool(record.get("comprobante_detectado") or det_status == "cdp_detectado")
        id_completa = bool(record.get("identidad_completa"))
        id_status = str(
            record.get("identity_status")
            or ((IDENTITY_COMPLETE if id_completa else IDENTITY_INCOMPLETE) if is_cdp else IDENTITY_NOT_APPLICABLE)
        )
        # La clave fiscal candidata puede llegar como tupla; se guarda como texto unido con guiones.
        cand_key = record.get("candidate_fiscal_key")
        if isinstance(cand_key, (tuple, list)):
            cand_key = "-".join(str(x) for x in cand_key)
        elif cand_key is not None:
            cand_key = str(cand_key)
        # Paso 2: ordenar los valores en el mismo orden que las columnas del INSERT.
        # Las listas y los diccionarios se guardan como JSON.
        values = (
            record.get("source_id"),
            original_path,
            json.dumps(container_chain, ensure_ascii=False),
            member_path,
            int(record.get("nested_depth") or 0),
            display_location,
            str(record.get("source_format") or "unknown"),
            str(record.get("reading_source") or "native_pdf_text"),
            int(record.get("page_count") or 0),
            int(record.get("page_start") or 1),
            int(record.get("page_end") or 1),
            int(record.get("block_start") or 1),
            int(record.get("block_end") or 1),
            det_status,
            id_status,
            int(is_cdp),
            int(id_completa),
            int(bool(record.get("renombrable"))),
            str(record.get("classification_rule") or ""),
            record.get("non_receipt_type"),
            str(record.get("decision_comment") or ""),
            str(record.get("detection_reason") or ""),
            record.get("detection_excerpt"),
            str(record.get("identity_reason") or ""),
            record.get("identity_excerpt"),
            int(bool(record.get("recipient_ruc_match"))),
            str(record.get("recipient_role_status") or "not_found"),
            str(record.get("issuer_role_status") or "not_found"),
            json.dumps(list(record.get("third_party_rucs") or []), ensure_ascii=False),
            cand_key,
            json.dumps(list(record.get("secondary_review_fields") or []), ensure_ascii=False),
            record.get("general_description"),
            int(bool(record.get("requires_human_review", True))),
            json.dumps(line_items, ensure_ascii=False),
            json.dumps(record.get("derived_checks") or [], ensure_ascii=False),
            json.dumps(record.get("field_statuses") or {}, ensure_ascii=False),
            json.dumps(record.get("fields_original") or {}, ensure_ascii=False),
            json.dumps(record.get("fields_normalized") or {}, ensure_ascii=False),
            json.dumps(record.get("field_evidence") or {}, ensure_ascii=False),
            json.dumps(record.get("confidence") or {}, ensure_ascii=False),
            json.dumps(record.get("candidates") or {}, ensure_ascii=False),
            json.dumps(record.get("observations") or [], ensure_ascii=False),
            int(bool(record.get("read_error"))),
            json.dumps(record.get("read_error_details") or [], ensure_ascii=False),
            str(record.get("extracted_text") or ""),
            record.get("text_sidecar_path"),
            record.get("normalized_filename"),
            record.get("organized_path"),
            now,
        )
        with self._connection() as connection:
            # Paso 3: insertar el documento y, si trae detalle de productos o servicios, insertar también cada fila.
            cursor = connection.execute(
                """INSERT INTO documents (
                    source_id, original_path, container_chain_json, member_path,
                    nested_depth, display_location, source_format, reading_source,
                    page_count, page_start, page_end, block_start, block_end,
                    detection_status, identity_status, comprobante_detectado, identidad_completa,
                    renombrable, classification_rule, non_receipt_type, decision_comment,
                    detection_reason, detection_excerpt, identity_reason, identity_excerpt,
                    recipient_ruc_match, recipient_role_status, issuer_role_status,
                    third_party_rucs_json, candidate_fiscal_key, secondary_review_fields_json,
                    general_description,
                    requires_human_review, line_items_json, derived_checks_json,
                    field_statuses_json, fields_original_json, fields_normalized_json,
                    field_evidence_json, confidence_json, candidates_json,
                    observations_json, read_error, read_error_details_json,
                    extracted_text, text_sidecar_path, normalized_filename,
                    organized_path, created_at
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                values,
            )
            doc_id = int(cursor.lastrowid)
            if line_items:
                item_rows = [
                    (
                        doc_id,
                        int(it.get("item_index") or idx),
                        it.get("code"),
                        it.get("description_observed"),
                        it.get("description_normalized"),
                        it.get("quantity_observed"),
                        it.get("quantity_normalized"),
                        it.get("unit_observed"),
                        it.get("unit_normalized"),
                        it.get("unit_price_observed"),
                        it.get("unit_price_normalized"),
                        it.get("line_total_observed"),
                        it.get("line_total_normalized"),
                        int(it.get("page_number") or record.get("page_start") or 1),
                        int(it.get("block_number") or record.get("block_start") or 1),
                        it.get("line_number"),
                        json.dumps(it.get("bbox"), ensure_ascii=False) if it.get("bbox") is not None else None,
                        str(it.get("excerpt") or ""),
                        str(it.get("source") or record.get("reading_source") or "native_pdf_text"),
                        it.get("ocr_confidence"),
                        it.get("heuristic_confidence"),
                        it.get("missing_reason"),
                    )
                    for idx, it in enumerate(line_items, start=1)
                ]
                connection.executemany(
                    """INSERT INTO document_line_items (
                        document_id, item_index, code, description_observed,
                        description_normalized, quantity_observed, quantity_normalized,
                        unit_observed, unit_normalized, unit_price_observed,
                        unit_price_normalized, line_total_observed, line_total_normalized,
                        page_number, block_number, line_number, bbox_json, excerpt,
                        source, ocr_confidence, heuristic_confidence, missing_reason
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    item_rows,
                )
        # Paso 4: se aplican de una vez los datos que se calculan después (padrón, verificación, reconciliación).
        self.update_document_post_processing(doc_id, record)
        return doc_id

    # Actualiza campos de un documento ya guardado. Sólo deja pasar una lista cerrada de columnas,
    # así no se puede escribir en otras por error.
    def update_document_post_processing(self, document_id: int, updates: Mapping[str, Any]) -> None:
        """Actualiza estado de padrón, validez oficial, reconciliación de clave y copia organizada tras conciliar el lote."""
        if not updates:
            return
        allowed_scalar = {
            "identity_status",
            "identidad_completa",
            "renombrable",
            "identity_reason",
            "candidate_fiscal_key",
            "normalized_filename",
            "organized_path",
            "official_validation_status",
            "official_validation_date",
            "official_validation_source",
            "official_validation_reference",
            "official_validation_reason",
            "official_validation_channel",
            "emitter_padron_status",
            "emitter_padron_razon_social",
            "emitter_padron_estado",
            "emitter_padron_cond_domicilio",
            "emitter_padron_cutoff_date",
            "emitter_padron_source",
            "emitter_name_discrepancy",
            "recipient_padron_status",
            "recipient_padron_razon_social",
            "recipient_padron_estado",
            "recipient_padron_cond_domicilio",
            "key_reconciliation_status",
            "key_reconciliation_reason",
            "is_primary_representation",
            "eligible_for_valid_sheet",
            "requires_human_review",
        }
        assignments: list[str] = []
        params: list[Any] = []
        # Estas columnas son verdadero/falso y se guardan como 0 o 1.
        bool_cols = {
            "identidad_completa",
            "renombrable",
            "emitter_name_discrepancy",
            "is_primary_representation",
            "eligible_for_valid_sheet",
            "requires_human_review",
        }
        # Armamos el UPDATE sólo con lo que venga permitido; lo demás se ignora.
        for key, val in updates.items():
            if key in allowed_scalar:
                assignments.append(f"{key} = ?")
                params.append(int(bool(val)) if key in bool_cols else val)
            elif key == "observations":
                assignments.append("observations_json = ?")
                params.append(json.dumps(list(val or []), ensure_ascii=False))
            elif key == "secondary_review_fields":
                assignments.append("secondary_review_fields_json = ?")
                params.append(json.dumps(list(val or []), ensure_ascii=False))
        if not assignments:
            return
        params.append(int(document_id))
        sql = f"UPDATE documents SET {', '.join(assignments)} WHERE id = ?"
        with self._connection() as connection:
            connection.execute(sql, params)

    # DEPRECADO: nadie lo llama; los documentos se leen todos juntos con list_documents.
    def get_document(self, document_id: int) -> dict[str, Any] | None:
        with self._connection() as connection:
            row = connection.execute("SELECT * FROM documents WHERE id=?", (document_id,)).fetchone()
        return None if row is None else self._decode_row(row)

    # Devuelve todos los documentos de la corrida, o sólo los de un estado de detección.
    def list_documents(self, detection_status: str | None = None) -> list[dict[str, Any]]:
        with self._connection() as connection:
            rows = connection.execute("SELECT * FROM documents ORDER BY id").fetchall()
        decoded = [self._decode_row(row) for row in rows]
        if detection_status is None:
            return decoded
        target_status = normalize_detection_status(detection_status)
        return [doc for doc in decoded if doc.get("detection_status") == target_status]

    # DEPRECADO: nadie lo llama; el detalle viaja dentro de cada documento.
    def list_line_items(self, document_id: int | None = None) -> list[dict[str, Any]]:
        with self._connection() as connection:
            if document_id is None:
                rows = connection.execute("SELECT * FROM document_line_items ORDER BY document_id, item_index").fetchall()
            else:
                rows = connection.execute(
                    "SELECT * FROM document_line_items WHERE document_id=? ORDER BY item_index",
                    (document_id,),
                ).fetchall()
        result: list[dict[str, Any]] = []
        for row in rows:
            d = dict(row)
            raw_bbox = d.pop("bbox_json", None)
            d["bbox"] = json.loads(raw_bbox) if raw_bbox else None
            result.append(d)
        return result

    # Guarda los eventos del proceso (qué etapa se ejecutó, cómo salió y cuánto tardó) asociados a un documento.
    def add_events(self, document_id: int | None, events: list[StageEvent]) -> None:
        if not events:
            return
        now = _utc_now()
        values = [
            (
                document_id,
                event.source_id,
                event.member_path,
                json.dumps(list(event.container_chain), ensure_ascii=False),
                event.stage,
                event.status,
                event.message,
                event.page_number,
                event.duration_ms,
                json.dumps(event.details or {}, ensure_ascii=False),
                now,
            )
            for event in events
        ]
        with self._connection() as connection:
            connection.executemany(
                """INSERT INTO processing_events(
                    document_id, source_id, member_path, container_chain_json,
                    stage, status, message, page_number, duration_ms, details_json, created_at
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                values,
            )

    # DEPRECADO: nadie lo llama.
    def get_events(self, document_id: int) -> list[dict[str, Any]]:
        with self._connection() as connection:
            rows = connection.execute(
                "SELECT * FROM processing_events WHERE document_id=? ORDER BY id", (document_id,)
            ).fetchall()
        result: list[dict[str, Any]] = []
        for row in rows:
            d = dict(row)
            d["container_chain"] = json.loads(d.pop("container_chain_json") or "[]")
            d["details"] = json.loads(d.pop("details_json") or "{}")
            result.append(d)
        return result

    # DEPRECADO: nadie lo llama.
    def list_all_events(self) -> list[dict[str, Any]]:
        with self._connection() as connection:
            rows = connection.execute("SELECT * FROM processing_events ORDER BY id").fetchall()
        result: list[dict[str, Any]] = []
        for row in rows:
            d = dict(row)
            d["container_chain"] = json.loads(d.pop("container_chain_json") or "[]")
            d["details"] = json.loads(d.pop("details_json") or "{}")
            result.append(d)
        return result
