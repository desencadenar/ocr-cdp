from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
import io
import json
from pathlib import Path
import re
import sqlite3
from typing import Any, Iterable, Mapping, Sequence
import zipfile

from .tax import validate_ruc



# Estados posibles al buscar un RUC en el padrón local: encontrado, no encontrado, sin padrón,
# instantánea demasiado vieja, archivo inválido o RUC mal formado.
PADRON_PRESENTE = "presente"
PADRON_AUSENTE = "ausente"
PADRON_NO_DISPONIBLE = "padron_no_disponible"
PADRON_OBSOLETO = "obsoleto"
PADRON_ARCHIVO_INVALIDO = "archivo_invalido"
PADRON_FORMATO_INVALIDO = "formato_invalido"

# SQLite limita los parámetros por consulta, así que se busca de a 900 RUC como máximo.
_BATCH_SQL_PARAM_LIMIT = 900
# Para comparar nombres se dejan sólo letras y números.
_NAME_CLEAN_RE = re.compile(r"[^A-Z0-9ÁÉÍÓÚÑ]+")
# Formas societarias que se ignoran al comparar nombres (S.A.C., S.R.L., E.I.R.L., etc.).
_CORP_SUFFIX_RE = re.compile(
    r"\b(S\.?A\.?C\.?|S\.?A\.?A\.?|S\.?A\.?|S\.?R\.?L\.?|E\.?I\.?R\.?L\.?|SUCURSAL DEL PERU|SOCIEDAD ANONIMA CERRADA|SOCIEDAD ANONIMA)\b"
)


# Deja un nombre en forma comparable: mayúsculas, sin tildes, sin forma societaria y sin puntuación.
def _normalize_company_name(value: str | None) -> str:
    if not value:
        return ""
    upper = value.upper().strip()
    upper = (
        upper.replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
    )
    upper = _CORP_SUFFIX_RE.sub(" ", upper)
    tokens = [tok for tok in _NAME_CLEAN_RE.split(upper) if tok]
    return " ".join(tokens)


# ¿El nombre leído en el comprobante contradice al del padrón? Sólo da True si ambos existen y son distintos.
# Un nombre abreviado que está contenido en el otro no cuenta como discrepancia.
def names_disagree(observed_name: str | None, padron_name: str | None) -> bool:
    """Devuelve True solo cuando ambos nombres existen y difieren tras normalizar puntuación y sufijos societarios."""
    norm_obs = _normalize_company_name(observed_name)
    norm_pad = _normalize_company_name(padron_name)
    if not norm_obs or not norm_pad:
        return False
    if norm_obs == norm_pad:
        return False
    # Permitir coincidencia por prefijo/subcadena cuando el comprobante abrevia la razón social larga
    if len(norm_obs) >= 6 and (norm_obs in norm_pad or norm_pad in norm_obs):
        return False
    return True


# Resultado de buscar un RUC: su estado en el padrón (encontrado, ausente...) y, si lo hay,
# su razón social, su estado y su condición de domicilio, junto con la fecha de corte del padrón.
@dataclass(slots=True)
class PadronLookupEntry:
    ruc: str
    status: str
    valid_format: bool
    valid_checksum: bool
    razon_social: str | None = None
    estado: str | None = None
    cond_domicilio: str | None = None
    ubigeo: str | None = None
    cutoff_date: str | None = None
    source_path: str | None = None
    reason: str | None = None

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


# Resultado de una búsqueda por lote: en qué estado estaba la base, cuántos RUC se consultaron
# y qué se encontró de cada uno.
@dataclass(slots=True)
class PadronBatchReport:
    db_status: str
    db_path: str | None
    cutoff_date: str | None
    source_path: str | None
    queried_count: int = 0
    present_count: int = 0
    absent_count: int = 0
    obsolete_count: int = 0
    invalid_ruc_count: int = 0
    entries: dict[str, PadronLookupEntry] = field(default_factory=dict)
    reason: str | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "db_status": self.db_status,
            "db_path": self.db_path,
            "cutoff_date": self.cutoff_date,
            "source_path": self.source_path,
            "queried_count": self.queried_count,
            "present_count": self.present_count,
            "absent_count": self.absent_count,
            "obsolete_count": self.obsolete_count,
            "invalid_ruc_count": self.invalid_ruc_count,
            "reason": self.reason,
            "entries": {k: v.as_dict() for k, v in self.entries.items()},
        }


# DEPRECADO: sólo lo devuelve cargar_padron_staging(), que nadie llama.
@dataclass(slots=True)
class PadronStagingResult:
    db_path: str
    source_path: str
    cutoff_date: str
    processed_rows: int
    unique_rucs: int
    rejected_rows: int
    duplicate_identical_rows: int
    conflict_rows: int
    conflicts: list[dict[str, Any]] = field(default_factory=list)
    rejections: list[dict[str, Any]] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


# Consulta de sólo lectura sobre el padrón RUC guardado en una base SQLite local. No usa internet.
# Sirve únicamente para contrastar un RUC con su nombre registrado y la fecha de corte del padrón;
# no se usa para inventar nombres observados, roles ni fechas.
class LocalPadronAdapter:
    """Adaptador local de solo lectura para consultar el Padrón RUC en SQLite sin dependencias de red."""

    # Guarda la ruta de la base, la antigüedad máxima aceptada y la fecha de referencia (hoy, por omisión).
    def __init__(
        self,
        db_path: str | Path | None = None,
        *,
        max_age_days: int | None = None,
        reference_date: datetime | None = None,
    ) -> None:
        self.db_path = Path(db_path) if db_path is not None else None
        self.max_age_days = max_age_days
        self.reference_date = reference_date or datetime.now(timezone.utc)

    # Revisa la base antes de usarla: que exista, que sea SQLite, que tenga la tabla padron con las columnas esperadas
    # y que no esté demasiado vieja. Devuelve su estado y la fecha de corte.
    def _inspect_db(self) -> tuple[str, str | None, str | None, str | None]:
        """Devuelve (db_status, cutoff_date, source_path, reason)."""
        # Paso 1: ¿hay una ruta configurada y existe el archivo?
        if self.db_path is None:
            return PADRON_NO_DISPONIBLE, None, None, "Ruta de base del padrón RUC no configurada."
        if not self.db_path.exists() or not self.db_path.is_file():
            return (
                PADRON_NO_DISPONIBLE,
                None,
                str(self.db_path),
                f"No existe el archivo SQLite del padrón RUC en {self.db_path}.",
            )
        try:
            # Paso 2: ¿es un SQLite válido, con la tabla y las columnas necesarias? Se abre en modo de sólo lectura.
            stat = self.db_path.stat()
            if stat.st_size < 100:
                return (
                    PADRON_ARCHIVO_INVALIDO,
                    None,
                    str(self.db_path),
                    f"Archivo del padrón vacío o truncado ({stat.st_size} bytes).",
                )
            with self.db_path.open("rb") as fh:
                header = fh.read(16)
            if not header.startswith(b"SQLite format 3\x00"):
                return (
                    PADRON_ARCHIVO_INVALIDO,
                    None,
                    str(self.db_path),
                    "El archivo del padrón no tiene cabecera SQLite válida.",
                )
            uri = f"file:{self.db_path.resolve().as_posix()}?mode=ro"
            with sqlite3.connect(uri, uri=True, timeout=10) as conn:
                conn.execute("PRAGMA query_only = ON")
                tables = {
                    row[0]
                    for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
                }
                if "padron" not in tables:
                    return (
                        PADRON_ARCHIVO_INVALIDO,
                        None,
                        str(self.db_path),
                        "La base SQLite no contiene la tabla requerida 'padron'.",
                    )
                cols = {row[1] for row in conn.execute("PRAGMA table_info(padron)").fetchall()}
                required_cols = {"ruc", "razon_social", "estado", "cond_domicilio"}
                if not required_cols.issubset(cols):
                    return (
                        PADRON_ARCHIVO_INVALIDO,
                        None,
                        str(self.db_path),
                        f"La tabla 'padron' no contiene las columnas esperadas ({sorted(required_cols)}).",
                    )
                # Paso 3: la fecha de corte sale de la tabla padron_metadata; si no existe, de la fecha de modificación del archivo.
                cutoff_date: str | None = None
                source_path = str(self.db_path)
                if "padron_metadata" in tables:
                    meta_rows = dict(conn.execute("SELECT key, value FROM padron_metadata").fetchall())
                    cutoff_date = meta_rows.get("cutoff_date")
                    source_path = meta_rows.get("source_path") or source_path
                if not cutoff_date:
                    cutoff_date = datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).date().isoformat()
        except sqlite3.Error as exc:
            return (
                PADRON_ARCHIVO_INVALIDO,
                None,
                str(self.db_path),
                f"Error SQLite al inspeccionar el padrón: {exc}",
            )
        except OSError as exc:
            return (
                PADRON_NO_DISPONIBLE,
                None,
                str(self.db_path),
                f"No se pudo abrir el archivo del padrón: {exc}",
            )

        # Paso 4: si se fijó una antigüedad máxima y la instantánea es más vieja, se marca como obsoleta.
        if self.max_age_days is not None and cutoff_date:
            try:
                cutoff_dt = datetime.fromisoformat(cutoff_date)
                if cutoff_dt.tzinfo is None:
                    cutoff_dt = cutoff_dt.replace(tzinfo=timezone.utc)
                age_days = (self.reference_date - cutoff_dt).days
                if age_days > self.max_age_days:
                    return (
                        PADRON_OBSOLETO,
                        cutoff_date,
                        source_path,
                        f"Instantánea del padrón con fecha {cutoff_date} supera la antigüedad máxima ({age_days} > {self.max_age_days} días).",
                    )
            except ValueError:
                pass

        return "disponible", cutoff_date, source_path, None

    # DEPRECADO: nadie lo llama. Imitaba la interfaz de otra herramienta externa de consulta de RUC.
    def consultar(self, ruc: str) -> dict[str, Any] | None:
        """Contrato compatible con una herramienta externa de consulta de RUC (método consultar)."""
        lote = self.consultar_lote([ruc])
        return lote.get(str(ruc).strip())

    # DEPRECADO: sólo lo usaba consultar(). El programa llama directamente a evaluate_rucs().
    def consultar_lote(self, rucs: Sequence[str]) -> dict[str, dict[str, Any]]:
        """Contrato compatible con PadronRUC.consultar_lote sin cargar dependencias externas."""
        report = self.evaluate_rucs(rucs)
        out: dict[str, dict[str, Any]] = {}
        for ruc, entry in report.entries.items():
            if entry.razon_social is not None or entry.estado is not None:
                out[ruc] = {
                    "ruc": entry.ruc,
                    "razon_social": entry.razon_social or "",
                    "estado": entry.estado or "",
                    "cond_domicilio": entry.cond_domicilio or "",
                    "ubigeo": entry.ubigeo or "",
                    "fuente": "PADRON_LOCAL",
                    "status": entry.status,
                    "cutoff_date": entry.cutoff_date,
                }
        return out

    # El punto de entrada que sí se usa: busca una lista de RUC en el padrón y devuelve un informe
    # con el resultado de cada uno.
    def evaluate_rucs(self, rucs: Iterable[str | None]) -> PadronBatchReport:
        # Paso 1: ver en qué estado está la base.
        db_status, cutoff_date, source_path, db_reason = self._inspect_db()
        report = PadronBatchReport(
            db_status=db_status,
            db_path=str(self.db_path) if self.db_path is not None else None,
            cutoff_date=cutoff_date,
            source_path=source_path,
            reason=db_reason,
        )

        # Paso 2: limpiar la lista de RUC (sin vacíos ni repetidos).
        cleaned_unique: list[str] = []
        seen: set[str] = set()
        for raw in rucs:
            if raw is None:
                continue
            text = str(raw).strip()
            if not text or text in seen:
                continue
            seen.add(text)
            cleaned_unique.append(text)

        # Paso 3: separar los RUC mal formados o con dígito de control incorrecto. Esos no se buscan.
        valid_to_query: list[str] = []
        for candidate in cleaned_unique:
            validation = validate_ruc(candidate)
            if not (validation.format_valid and validation.checksum_valid):
                report.invalid_ruc_count += 1
                report.entries[candidate] = PadronLookupEntry(
                    ruc=candidate,
                    status=PADRON_FORMATO_INVALIDO,
                    valid_format=bool(validation.format_valid),
                    valid_checksum=bool(validation.checksum_valid),
                    cutoff_date=cutoff_date,
                    source_path=source_path,
                    reason=(
                        "dígito verificador inválido"
                        if validation.format_valid
                        else "formato de RUC inválido"
                    ),
                )
            else:
                valid_to_query.append(candidate)

        # Paso 4: si no hay una base utilizable, todos los RUC quedan con el estado de la base y no se consulta nada.
        report.queried_count = len(valid_to_query)
        if not valid_to_query:
            return report

        if db_status in {PADRON_NO_DISPONIBLE, PADRON_ARCHIVO_INVALIDO}:
            for candidate in valid_to_query:
                report.entries[candidate] = PadronLookupEntry(
                    ruc=candidate,
                    status=db_status,
                    valid_format=True,
                    valid_checksum=True,
                    cutoff_date=cutoff_date,
                    source_path=source_path,
                    reason=db_reason,
                )
            return report

        assert self.db_path is not None
        # Paso 5: consultar la base por lotes de hasta 900 RUC.
        found_rows: dict[str, tuple[str, str, str, str]] = {}
        uri = f"file:{self.db_path.resolve().as_posix()}?mode=ro"
        with sqlite3.connect(uri, uri=True, timeout=15) as conn:
            conn.execute("PRAGMA query_only = ON")
            cols = {row[1] for row in conn.execute("PRAGMA table_info(padron)").fetchall()}
            has_ubigeo = "ubigeo" in cols
            for offset in range(0, len(valid_to_query), _BATCH_SQL_PARAM_LIMIT):
                batch = valid_to_query[offset : offset + _BATCH_SQL_PARAM_LIMIT]
                placeholders = ",".join("?" * len(batch))
                if has_ubigeo:
                    sql = f"SELECT ruc, razon_social, estado, cond_domicilio, ubigeo FROM padron WHERE ruc IN ({placeholders})"
                else:
                    sql = f"SELECT ruc, razon_social, estado, cond_domicilio, '' AS ubigeo FROM padron WHERE ruc IN ({placeholders})"
                for ruc_val, razon, estado, cond, ubigeo in conn.execute(sql, batch).fetchall():
                    found_rows[str(ruc_val)] = (
                        str(razon or "").strip(),
                        str(estado or "").strip(),
                        str(cond or "").strip(),
                        str(ubigeo or "").strip(),
                    )

        # Paso 6: para cada RUC, anotar si apareció (presente), si no apareció (ausente)
        # o si la base es demasiado vieja (obsoleto).
        for candidate in valid_to_query:
            if candidate in found_rows:
                razon, estado, cond, ubigeo = found_rows[candidate]
                entry_status = PADRON_OBSOLETO if db_status == PADRON_OBSOLETO else PADRON_PRESENTE
                if entry_status == PADRON_OBSOLETO:
                    report.obsolete_count += 1
                else:
                    report.present_count += 1
                report.entries[candidate] = PadronLookupEntry(
                    ruc=candidate,
                    status=entry_status,
                    valid_format=True,
                    valid_checksum=True,
                    razon_social=razon,
                    estado=estado,
                    cond_domicilio=cond,
                    ubigeo=ubigeo,
                    cutoff_date=cutoff_date,
                    source_path=source_path,
                    reason=db_reason if entry_status == PADRON_OBSOLETO else "RUC encontrado en padrón local",
                )
            else:
                entry_status = PADRON_OBSOLETO if db_status == PADRON_OBSOLETO else PADRON_AUSENTE
                if entry_status == PADRON_OBSOLETO:
                    report.obsolete_count += 1
                else:
                    report.absent_count += 1
                report.entries[candidate] = PadronLookupEntry(
                    ruc=candidate,
                    status=entry_status,
                    valid_format=True,
                    valid_checksum=True,
                    cutoff_date=cutoff_date,
                    source_path=source_path,
                    reason=(
                        db_reason
                        if entry_status == PADRON_OBSOLETO
                        else "RUC con formato y dígito válidos ausente en la instantánea local del padrón"
                    ),
                )

        return report


# DEPRECADO: nadie lo llama desde el programa. Cargaba el Padrón Reducido de SUNAT (TXT o ZIP) en una base SQLite.
# La base del padrón se prepara fuera de este programa.
def cargar_padron_staging(
    source_path: str | Path,
    db_path: str | Path,
    *,
    cutoff_date: str | None = None,
    encoding: str = "cp1252",
) -> PadronStagingResult:
    """Carga un archivo TXT o ZIP del Padrón Reducido SUNAT en SQLite mediante tabla staging.

    A diferencia de INSERT OR REPLACE silencioso, contabiliza filas procesadas, RUCs únicos,
    filas rechazadas por formato/dígito verificador y conflictos cuando un mismo RUC aparece
    con atributos distintos en el archivo de entrada.
    """
    src = Path(source_path)
    dst = Path(db_path)
    if not src.exists():
        raise FileNotFoundError(f"No existe el archivo fuente del padrón: {src}")

    effective_cutoff = cutoff_date or datetime.fromtimestamp(src.stat().st_mtime, tz=timezone.utc).date().isoformat()

    def _open_text_stream() -> io.TextIOWrapper:
        if src.suffix.lower() == ".zip":
            zf = zipfile.ZipFile(src, "r")
            txt_names = [name for name in zf.namelist() if name.lower().endswith(".txt")]
            if not txt_names:
                zf.close()
                raise ValueError(f"El archivo ZIP {src} no contiene ningún miembro .txt")
            raw_stream = zf.open(txt_names[0], "r")
            return io.TextIOWrapper(raw_stream, encoding=encoding, errors="replace")
        return src.open("r", encoding=encoding, errors="replace")

    dst.parent.mkdir(parents=True, exist_ok=True)
    processed_rows = 0
    unique_rucs = 0
    rejected_rows = 0
    duplicate_identical_rows = 0
    conflict_rows = 0
    conflicts: list[dict[str, Any]] = []
    rejections: list[dict[str, Any]] = []

    with sqlite3.connect(dst, timeout=30) as conn:
        conn.execute("PRAGMA synchronous = NORMAL")
        conn.executescript(
            """
            DROP TABLE IF EXISTS padron_staging;
            CREATE TABLE padron_staging (
                ruc TEXT PRIMARY KEY,
                razon_social TEXT NOT NULL,
                estado TEXT NOT NULL,
                cond_domicilio TEXT NOT NULL,
                ubigeo TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS padron_conflicts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                line_number INTEGER NOT NULL,
                ruc TEXT NOT NULL,
                existing_json TEXT NOT NULL,
                conflicting_json TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            DELETE FROM padron_conflicts;
            """
        )

        with _open_text_stream() as text_stream:
            header_line = text_stream.readline()
            if not header_line:
                raise ValueError("El archivo del padrón está vacío.")
            header_upper = header_line.upper()
            if "RUC" not in header_upper or "|" not in header_line:
                raise ValueError(f"Encabezado del padrón inválido: {header_line.strip()!r}")

            now_iso = datetime.now(timezone.utc).isoformat()
            for line_number, raw_line in enumerate(text_stream, start=2):
                stripped = raw_line.strip()
                if not stripped:
                    continue
                processed_rows += 1
                parts = raw_line.rstrip("\r\n").split("|")
                if len(parts) < 4:
                    rejected_rows += 1
                    if len(rejections) < 100:
                        rejections.append(
                            {"line_number": line_number, "ruc": parts[0] if parts else "", "reason": "columnas_insuficientes"}
                        )
                    continue
                ruc = parts[0].strip()
                validation = validate_ruc(ruc)
                if not (validation.format_valid and validation.checksum_valid):
                    rejected_rows += 1
                    if len(rejections) < 100:
                        rejections.append(
                            {
                                "line_number": line_number,
                                "ruc": ruc,
                                "reason": "digito_verificador_invalido" if validation.format_valid else "formato_ruc_invalido",
                            }
                        )
                    continue

                razon = parts[1].strip()
                estado = parts[2].strip()
                cond = parts[3].strip()
                ubigeo = parts[4].strip() if len(parts) > 4 else ""

                existing = conn.execute(
                    "SELECT razon_social, estado, cond_domicilio, ubigeo FROM padron_staging WHERE ruc = ?",
                    (ruc,),
                ).fetchone()
                if existing is None:
                    conn.execute(
                        "INSERT INTO padron_staging (ruc, razon_social, estado, cond_domicilio, ubigeo) VALUES (?,?,?,?,?)",
                        (ruc, razon, estado, cond, ubigeo),
                    )
                    unique_rucs += 1
                else:
                    ex_tuple = (str(existing[0]), str(existing[1]), str(existing[2]), str(existing[3]))
                    new_tuple = (razon, estado, cond, ubigeo)
                    if ex_tuple == new_tuple:
                        duplicate_identical_rows += 1
                    else:
                        conflict_rows += 1
                        conflict_item = {
                            "line_number": line_number,
                            "ruc": ruc,
                            "existing": {
                                "razon_social": ex_tuple[0],
                                "estado": ex_tuple[1],
                                "cond_domicilio": ex_tuple[2],
                                "ubigeo": ex_tuple[3],
                            },
                            "conflicting": {
                                "razon_social": razon,
                                "estado": estado,
                                "cond_domicilio": cond,
                                "ubigeo": ubigeo,
                            },
                        }
                        if len(conflicts) < 100:
                            conflicts.append(conflict_item)
                        conn.execute(
                            "INSERT INTO padron_conflicts (line_number, ruc, existing_json, conflicting_json, created_at) VALUES (?,?,?,?,?)",
                            (
                                line_number,
                                ruc,
                                json.dumps(conflict_item["existing"], ensure_ascii=False),
                                json.dumps(conflict_item["conflicting"], ensure_ascii=False),
                                now_iso,
                            ),
                        )

        conn.executescript(
            """
            DROP TABLE IF EXISTS padron;
            ALTER TABLE padron_staging RENAME TO padron;
            CREATE INDEX IF NOT EXISTS idx_padron_estado ON padron(estado);
            CREATE TABLE IF NOT EXISTS padron_metadata (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            """
        )
        metadata_pairs = [
            ("source_path", str(src.resolve())),
            ("cutoff_date", effective_cutoff),
            ("loaded_at", datetime.now(timezone.utc).isoformat()),
            ("processed_rows", str(processed_rows)),
            ("unique_rucs", str(unique_rucs)),
            ("rejected_rows", str(rejected_rows)),
            ("duplicate_identical_rows", str(duplicate_identical_rows)),
            ("conflict_rows", str(conflict_rows)),
        ]
        conn.executemany(
            "INSERT OR REPLACE INTO padron_metadata (key, value) VALUES (?, ?)",
            metadata_pairs,
        )
        conn.commit()

    return PadronStagingResult(
        db_path=str(dst),
        source_path=str(src),
        cutoff_date=effective_cutoff,
        processed_rows=processed_rows,
        unique_rucs=unique_rucs,
        rejected_rows=rejected_rows,
        duplicate_identical_rows=duplicate_identical_rows,
        conflict_rows=conflict_rows,
        conflicts=conflicts,
        rejections=rejections,
    )
