from __future__ import annotations

from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from dataclasses import dataclass, field, replace
from decimal import Decimal, InvalidOperation
import os
from pathlib import Path
import re
import threading
import time
from typing import Any, Callable, Mapping, Sequence

from .amount_panel import review_amount_panel
from .archives import ContainerWalker, WorkUnit
from .ciiu import ConsultaRucReport, load_consulta_ruc, refine_special_rate_check
from .config import AppConfig
from . import documents as _documents
from .documents import DocumentReadError, read_work_unit, reocr_line_bbox
from .reading_cache import CACHE_FILENAME, ReadingCache, reread_key, unit_key
from .extraction import parse_document
from .models import (
    AMBIGUOUS,
    DETECTED,
    IDENTITY_COMPLETE,
    IDENTITY_CONFLICT,
    IDENTITY_NOT_APPLICABLE,
    NOT_RECEIPT,
    READ_ERROR,
    ArchiveSummary,
    ExtractionResult,
    StageEvent,
)
from .ocr import OCREngine, TesseractOCREngine
from .padron import LocalPadronAdapter, compare_names_with_padron
from .pdf_organizer import organize_pdf_expediente
from .storage import StateStore
from .tax import normalize_correlativo, validate_ruc
# Un aviso de avance para la ventana: en qué unidad va, cuántas hay en total y los contadores acumulados.
@dataclass
class ProgressUpdate:
    stage: str
    current_unit: int
    total_units: int
    current_path: str
    detected_count: int
    ambiguous_count: int
    no_comprobante_count: int
    read_error_count: int
    zip_container_count: int
    message: str = ""


# Lo que quedó de cada documento procesado: su estado, sus datos, el resultado del padrón y de la reconciliación.
# Es el resumen que usa el resto del programa para exportar.
@dataclass
class DocumentOutcome:
    document_id: int
    original_path: str
    detection_status: str
    comprobante_detectado: bool
    identidad_completa: bool
    renombrable: bool
    page_start: int
    page_end: int
    block_start: int
    block_end: int
    normalized_filename: str | None
    output_path: str | None
    identity_status: str = IDENTITY_NOT_APPLICABLE
    source_id: int | None = None
    container_chain: list[str] = field(default_factory=list)
    member_path: str | None = None
    nested_depth: int = 0
    display_location: str = ""
    reading_source: str = "native_pdf_text"
    classification_rule: str = ""
    non_receipt_type: str | None = None
    decision_comment: str = ""
    detection_reason: str = ""
    detection_excerpt: str | None = None
    identity_reason: str = ""
    identity_excerpt: str | None = None
    recipient_ruc_match: bool = False
    recipient_role_status: str = "not_found"
    issuer_role_status: str = "not_found"
    third_party_rucs: list[str] = field(default_factory=list)
    candidate_fiscal_key: str | None = None
    secondary_review_fields: list[str] = field(default_factory=list)
    general_description: str | None = None
    requires_human_review: bool = True
    line_items: list[dict[str, Any]] = field(default_factory=list)
    derived_checks: list[dict[str, Any]] = field(default_factory=list)
    field_statuses: dict[str, str] = field(default_factory=dict)
    fields: dict[str, Any] = field(default_factory=dict)
    fields_original: dict[str, Any] = field(default_factory=dict)
    confidence: dict[str, float] = field(default_factory=dict)
    field_evidence: dict[str, Any] = field(default_factory=dict)
    candidates: dict[str, Any] = field(default_factory=dict)
    observations: list[str] = field(default_factory=list)
    read_error: bool = False
    read_error_details: list[str] = field(default_factory=list)
    text_sidecar_path: str | None = None
    # Verificación oficial. Este programa no la hace, así que queda «pendiente».
    official_validation_status: str = "pendiente"
    official_validation_date: str | None = None
    official_validation_source: str | None = None
    official_validation_reference: str | None = None
    official_validation_reason: str = ""
    official_validation_channel: str = "sin_canal"
    # Contraste con el padrón RUC (sólo si se configuró uno).
    emitter_padron_status: str = "no_aplica"
    emitter_padron_razon_social: str | None = None
    emitter_padron_estado: str | None = None
    emitter_padron_cond_domicilio: str | None = None
    emitter_padron_cutoff_date: str | None = None
    emitter_padron_source: str | None = None
    emitter_name_discrepancy: bool = False
    recipient_padron_status: str = "no_aplica"
    recipient_padron_razon_social: str | None = None
    recipient_padron_estado: str | None = None
    recipient_padron_cond_domicilio: str | None = None
    # Reconciliación de las copias del mismo comprobante.
    key_reconciliation_status: str = "sin_clave_completa"
    key_reconciliation_reason: str = ""
    is_primary_representation: bool = False
    eligible_for_valid_sheet: bool = False
    elapsed_seconds: float = 0.0

    # Alias del estado de detección, por compatibilidad con código antiguo.
    @property
    def status(self) -> str:
        return self.detection_status


# El resumen de toda la corrida: los resultados por documento, los contadores de archivos y ZIP,
# el estado final (completado o cancelado), los tiempos por etapa y los resultados del padrón y de la reconciliación.
@dataclass
class ScanSummary:
    outcomes: list[DocumentOutcome] = field(default_factory=list)
    discovered_unit_count: int = 0
    started_unit_count: int = 0
    processed_file_count: int = 0
    pending_unit_count: int = 0
    pending_units: list[str] = field(default_factory=list)
    excluded_xml_count: int = 0
    ignored_unsupported_count: int = 0
    top_level_file_count: int = 0
    zip_container_count: int = 0
    zip_member_count: int = 0
    cancelled: bool = False
    run_status: str = "completado"
    stage_timings_ms: dict[str, float] = field(default_factory=dict)
    archive_summaries: list[ArchiveSummary] = field(default_factory=list)
    padron_summary: dict[str, Any] = field(default_factory=dict)
    consulta_ruc_summary: dict[str, Any] = field(default_factory=dict)
    reconciliation_summary: dict[str, Any] = field(default_factory=dict)

    # Cuántos documentos hay de cada estado de detección.
    @property
    def counts(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for item in self.outcomes:
            counts[item.detection_status] = counts.get(item.detection_status, 0) + 1
        return counts

    # Cuántos documentos se detectaron como comprobante.
    @property
    def detected_count(self) -> int:
        return sum(item.detection_status == DETECTED for item in self.outcomes)

    # Cuántos quedaron indeterminados (sin contar los que fallaron al leerse).
    @property
    def ambiguous_count(self) -> int:
        return sum(item.detection_status == AMBIGUOUS and not item.read_error for item in self.outcomes)

    # Cuántos no son comprobante.
    @property
    def no_comprobante_count(self) -> int:
        return sum(item.detection_status == NOT_RECEIPT for item in self.outcomes)

    # Cuántos tuvieron error de lectura.
    @property
    def read_error_count(self) -> int:
        return sum(item.detection_status == READ_ERROR or item.read_error for item in self.outcomes)

    # El resumen en forma de diccionario, para mostrarlo o guardarlo.
    def as_dict(self) -> dict[str, Any]:
        return {
            "run_status": self.run_status,
            "discovered_unit_count": self.discovered_unit_count,
            "started_unit_count": self.started_unit_count,
            "processed_file_count": self.processed_file_count,
            "pending_unit_count": self.pending_unit_count,
            "pending_units": list(self.pending_units),
            "excluded_xml_count": self.excluded_xml_count,
            "ignored_unsupported_count": self.ignored_unsupported_count,
            "top_level_file_count": self.top_level_file_count,
            "zip_container_count": self.zip_container_count,
            "zip_member_count": self.zip_member_count,
            "cancelled": self.cancelled,
            "receipt_count": self.detected_count,
            "ambiguous_count": self.ambiguous_count,
            "no_comprobante_count": self.no_comprobante_count,
            "read_error_count": self.read_error_count,
            "counts": self.counts,
            "stage_timings_ms": {k: round(v, 3) for k, v in self.stage_timings_ms.items()},
            "archive_containers": [a.as_dict() for a in self.archive_summaries],
            "documents": [
                {
                    "document_id": item.document_id,
                    "source_id": item.source_id,
                    "original_path": item.original_path,
                    "container_chain": list(item.container_chain),
                    "member_path": item.member_path,
                    "nested_depth": item.nested_depth,
                    "display_location": item.display_location,
                    "reading_source": item.reading_source,
                    "detection_status": item.detection_status,
                    "identity_status": item.identity_status,
                    "classification_rule": item.classification_rule,
                    "non_receipt_type": item.non_receipt_type,
                    "comprobante_detectado": item.comprobante_detectado,
                    "identidad_completa": item.identidad_completa,
                    "renombrable": item.renombrable,
                    "recipient_ruc_match": item.recipient_ruc_match,
                    "recipient_role_status": item.recipient_role_status,
                    "issuer_role_status": item.issuer_role_status,
                    "third_party_rucs": list(item.third_party_rucs),
                    "candidate_fiscal_key": item.candidate_fiscal_key,
                    "secondary_review_fields": list(item.secondary_review_fields),
                    "requires_human_review": item.requires_human_review,
                    "decision_comment": item.decision_comment,
                    "detection_reason": item.detection_reason,
                    "identity_reason": item.identity_reason,
                    "page_start": item.page_start,
                    "page_end": item.page_end,
                    "block_start": item.block_start,
                    "block_end": item.block_end,
                    "normalized_filename": item.normalized_filename,
                    "output_path": item.output_path,
                    "text_sidecar_path": item.text_sidecar_path,
                    "official_validation_status": item.official_validation_status,
                    "official_validation_date": item.official_validation_date,
                    "official_validation_source": item.official_validation_source,
                    "official_validation_channel": item.official_validation_channel,
                    "emitter_padron_status": item.emitter_padron_status,
                    "key_reconciliation_status": item.key_reconciliation_status,
                    "is_primary_representation": item.is_primary_representation,
                    "eligible_for_valid_sheet": item.eligible_for_valid_sheet,
                    "read_error": item.read_error,
                    "elapsed_seconds": round(item.elapsed_seconds, 6),
                }
                for item in self.outcomes
            ],
            "padron_summary": dict(self.padron_summary),
            "consulta_ruc_summary": dict(self.consulta_ruc_summary),
            "reconciliation_summary": dict(self.reconciliation_summary),
        }


_QR_DATE_RE = re.compile(r"^(\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4})$")


def _parse_sunat_qr_payload(raw_qr: str) -> dict[str, str] | None:
    """Interpreta la cadena estándar SUNAT delimitada por `|` del código QR."""
    parts = [p.strip() for p in (raw_qr or "").split("|")]
    if len(parts) < 7:
        return None
    emitter_ruc = parts[0]
    if not (len(emitter_ruc) == 11 and emitter_ruc.isdigit()):
        return None
    doc_type = parts[1].zfill(2) if parts[1].isdigit() else parts[1]
    series = parts[2].upper()
    correlativo_res = normalize_correlativo(parts[3])
    igv = parts[4]
    total = parts[5]
    raw_date = parts[6]
    iso_date: str | None = None
    if _QR_DATE_RE.match(raw_date):
        if "-" in raw_date:
            iso_date = raw_date
        else:
            d, m, y = raw_date.split("/")
            iso_date = f"{y}-{m}-{d}"
    recipient_doc = parts[8] if len(parts) > 8 and parts[8].isdigit() else None
    return {
        "emitter_ruc": emitter_ruc,
        "document_type": doc_type,
        "series": series,
        "correlativo_normalized": correlativo_res.normalized or parts[3],
        "igv": igv,
        "total": total,
        "issue_date": iso_date or raw_date,
        "recipient_ruc": recipient_doc or "",
    }


def _qr_amount(raw: str) -> Decimal:
    """Importe del QR. El estándar usa punto decimal; se acepta también «18,00» (coma y dos decimales)."""
    value = raw.strip()
    if re.fullmatch(r"-?\d+,\d{1,2}", value):
        value = value.replace(",", ".")
    return Decimal(value.replace(",", ""))


def _crosscheck_sunat_qr(
    result: ExtractionResult,
    pages: Sequence[Any],
    *,
    qr_decoder_available: bool | None = None,
) -> None:
    """Coteja lo extraído visualmente contra el código QR SUNAT de las páginas del comprobante (si existe).

    Estados:
    - `coincide`: se compararon al menos RUC emisor, serie y total, y no hay diferencias.
    - `coincide_parcial`: no hay diferencias, pero faltó comparar alguno de esos tres campos.
    - `difiere`: algún campo comparado no coincide.
    - `qr_no_vinculado`: hay QR en la página, pero ninguno coincide en serie ni RUC emisor con lo leído.
    - `sin_qr`: se buscó y no hay QR; `qr_no_ejecutado`: no se buscó porque falta `zxingcpp`.
    """
    if not result.comprobante_detectado:
        return
    if qr_decoder_available is None:
        qr_decoder_available = _documents.QR_DECODER_AVAILABLE
    qr_entries: list[dict[str, str]] = []
    for p in pages:
        pno = getattr(p, "page_number", 1)
        if not (result.page_start <= pno <= result.page_end):
            continue
        for payload in getattr(p, "qr_payloads", ()) or ():
            parsed = _parse_sunat_qr_payload(payload)
            if parsed:
                qr_entries.append(parsed)

    if not qr_entries:
        if qr_decoder_available:
            result.derived_checks.append({"check": "sunat_qr_crosscheck", "status": "sin_qr", "consistent": None})
        else:
            result.derived_checks.append({
                "check": "sunat_qr_crosscheck",
                "status": "qr_no_ejecutado",
                "consistent": None,
                "reason": "No está instalado zxingcpp: el código QR no se buscó ni se cotejó.",
            })
        return

    # Se usa el QR que coincida en serie o RUC emisor con lo leído. Si ninguno coincide, no se toma
    # el primero de la página: puede ser de otro documento.
    qr = next(
        (
            entry for entry in qr_entries
            if (result.series and entry["series"] == result.series.upper())
            or (result.emitter_ruc and entry["emitter_ruc"] == result.emitter_ruc)
        ),
        None,
    )
    if qr is None:
        result.derived_checks.append({
            "check": "sunat_qr_crosscheck",
            "status": "qr_no_vinculado",
            "consistent": None,
            "qr_summary": "; ".join(
                f"{e['emitter_ruc']}-{e['document_type']}-{e['series']}-{e['correlativo_normalized']}" for e in qr_entries
            ),
            "reason": "Hay código QR en la página, pero ninguno coincide en serie ni en RUC emisor con lo leído del texto.",
        })
        result.observations.append("código QR en la página sin vínculo con la serie ni el RUC emisor leídos")
        if "qr_crosscheck" not in result.secondary_review_fields:
            result.secondary_review_fields.append("qr_crosscheck")
        return
    discrepancies: list[str] = []

    if result.emitter_ruc and qr["emitter_ruc"] != result.emitter_ruc:
        discrepancies.append(f"ruc_emisor(texto={result.emitter_ruc} vs qr={qr['emitter_ruc']})")
    if result.document_type and qr["document_type"] != result.document_type:
        discrepancies.append(f"tipo(texto={result.document_type} vs qr={qr['document_type']})")
    if result.series and qr["series"] != result.series.upper():
        discrepancies.append(f"serie(texto={result.series} vs qr={qr['series']})")
    if result.correlativo_normalized and qr["correlativo_normalized"] != result.correlativo_normalized:
        discrepancies.append(f"correlativo(texto={result.correlativo_normalized} vs qr={qr['correlativo_normalized']})")
    if result.issue_date and qr["issue_date"] and qr["issue_date"] != result.issue_date:
        discrepancies.append(f"fecha(texto={result.issue_date} vs qr={qr['issue_date']})")
    if result.igv and qr["igv"]:
        try:
            if abs(Decimal(result.igv) - _qr_amount(qr["igv"])) > Decimal("0.02"):
                discrepancies.append(f"igv(texto={result.igv} vs qr={qr['igv']})")
        except InvalidOperation:
            pass
    if result.total and qr["total"]:
        try:
            if abs(Decimal(result.total) - _qr_amount(qr["total"])) > Decimal("0.02"):
                discrepancies.append(f"total(texto={result.total} vs qr={qr['total']})")
        except InvalidOperation:
            pass
    if result.recipient_ruc and qr["recipient_ruc"] and len(qr["recipient_ruc"]) == 11 and qr["recipient_ruc"] != result.recipient_ruc:
        discrepancies.append(f"ruc_receptor(texto={result.recipient_ruc} vs qr={qr['recipient_ruc']})")

    not_compared = [
        label for label, text_value, qr_value in (
            ("ruc_emisor", result.emitter_ruc, qr["emitter_ruc"]),
            ("serie", result.series, qr["series"]),
            ("total", result.total, qr["total"]),
        )
        if not (text_value and qr_value)
    ]
    if discrepancies:
        status, consistent = "difiere", False
    elif not_compared:
        status, consistent = "coincide_parcial", None
    else:
        status, consistent = "coincide", True
    qr_check = {
        "check": "sunat_qr_crosscheck",
        "status": status,
        "consistent": consistent,
        "discrepancies": discrepancies,
        "qr_summary": f"{qr['emitter_ruc']}-{qr['document_type']}-{qr['series']}-{qr['correlativo_normalized']} (IGV={qr['igv']}, Total={qr['total']})",
    }
    if status == "coincide_parcial":
        qr_check["reason"] = f"Sin diferencias, pero no se pudo comparar: {', '.join(not_compared)}."
    result.derived_checks.append(qr_check)
    if discrepancies:
        result.observations.append(f"cotejo QR SUNAT difiere: {', '.join(discrepancies)}")
        if "qr_crosscheck" not in result.secondary_review_fields:
            result.secondary_review_fields.append("qr_crosscheck")


def _surgical_line_reocr(
    unit: WorkUnit,
    pages: Sequence[Any],
    results: Sequence[ExtractionResult],
    config: AppConfig,
    engine: OCREngine,
    reread_fn: Callable[[int, Any, int], str | None] | None = None,
) -> list[StageEvent]:
    """Relee con `--psm 7` solo las líneas puntuales que tuvieron error de dígito verificador o duda O/0.

    `reread_fn(página, bbox, rotación)` sustituye a Tesseract: graba las relecturas en la lectura guardada o las
    toma de ella en `redetectar`.
    """
    events: list[StageEvent] = []
    if reread_fn is None and not isinstance(engine, TesseractOCREngine):
        return events
    page_map = {getattr(p, "page_number", 1): p for p in pages}

    for res in results:
        if not res.comprobante_detectado:
            continue
        targets: list[tuple[str, Any]] = []
        # Caso 1: RUC emisor sin validar o serie marcada para revisión secundaria.
        for fname in ("emitter_ruc", "series", "correlativo", "total"):
            if fname in res.secondary_review_fields or (fname == "emitter_ruc" and not res.emitter_ruc):
                ev_list = res.field_evidence.get(fname) or []
                for ev in ev_list:
                    if ev.source == "ocr" and ev.bbox is not None:
                        targets.append((fname, ev))
                        break
        for fname, ev in targets[:3]:
            page = page_map.get(ev.page_number)
            rot_angle = int((getattr(page, "quality", {}) or {}).get("rotation_angle", 0)) if page else 0
            if reread_fn is not None:
                reread = reread_fn(ev.page_number, ev.bbox, rot_angle)
            else:
                reread = reocr_line_bbox(unit, ev.page_number, ev.bbox, config, engine, rotation_angle=rot_angle)
            if not reread:
                continue
            events.append(
                StageEvent(
                    "alternate_region_ocr",
                    "succeeded",
                    f"relectura de línea ({fname}): {reread[:80]}",
                    ev.page_number,
                )
            )
            if fname == "emitter_ruc" and not res.emitter_ruc:
                for tok in re.findall(r"(?<![0-9])[0-9]{11}(?![0-9])", reread):
                    val = validate_ruc(tok)
                    if val.format_valid and val.checksum_valid:
                        res.observations.append(
                            f"relectura de línea {fname} sugiere RUC válido {tok} (texto original: {ev.observed_value})"
                        )
                        res.candidates.setdefault("emitter_ruc", [])
                        if tok not in res.candidates["emitter_ruc"]:
                            res.candidates["emitter_ruc"].append(tok)
                        break
            elif reread.strip() != ev.excerpt.strip():
                res.observations.append(
                    f"relectura puntual de {fname}: «{reread}»"
                )
    return events


# El orquestador: junta todas las piezas (recorrer carpetas, leer, extraer, guardar y reconciliar).
# Una corrida corresponde a una carpeta de salida nueva.

def _apply_emitter_activity(item: "DocumentOutcome", consulta: ConsultaRucReport) -> None:
    """Anota el CIIU del emisor y, si la tasa especial estaba por verificar, la decide con él."""
    if item.detection_status != DETECTED:
        return
    em_ruc = str(item.fields.get("emitter_ruc") or item.fields_original.get("emitter_ruc") or "").strip()
    act = consulta.activities.get(em_ruc)
    if act is not None:
        item.derived_checks = [c for c in item.derived_checks if c.get("check") != "emitter_ciiu"]
        item.derived_checks.append({
            "check": "emitter_ciiu", "ruc": act.ruc, "principal": act.principal, "secondary": list(act.secondary),
            "estado": act.estado, "inicio_actividades": act.inicio_actividades, "source": act.source,
            "query_date": act.query_date, "label": act.label(),
        })
    for idx, chk in enumerate(item.derived_checks):
        if chk.get("check") != "igv_rate_check" or chk.get("status") != "tasa_especial_verificar":
            continue
        issue_date = item.fields.get("issue_date") or item.fields_original.get("issue_date")
        doc_type = item.fields.get("document_type") or item.fields_original.get("document_type")
        # Decisión del usuario (2026-11-20): con un CIIU Rev. 3 mixto deciden las descripciones del comprobante.
        item_texts = [str(li.get("description_normalized") or li.get("description_observed") or "") for li in item.line_items]
        descriptions = " | ".join(t for t in [item.general_description or "", *item_texts] if t)
        new = refine_special_rate_check(chk, act, issue_date=issue_date, document_type=doc_type, descriptions=descriptions)
        if new is chk:
            continue
        item.derived_checks[idx] = new
        if chk.get("reason") in item.observations:
            item.observations.remove(chk["reason"])
        item.observations.append(new["reason"])
        if new.get("status") == "tasa_especial_valida":
            if "igv_rate" in item.secondary_review_fields:
                item.secondary_review_fields.remove("igv_rate")
            # Mismo criterio que extraction.parse_document para «requiere revisión humana».
            item.requires_human_review = item.identity_status != IDENTITY_COMPLETE or bool(item.secondary_review_fields)

class FiscalProcessor:
    # Prepara la carpeta de salida (tiene que estar vacía o no existir), crea la base de estado
    # y guarda los datos de la corrida.
    def __init__(
        self,
        output_root: str | Path,
        config: AppConfig | None = None,
        engine: OCREngine | None = None,
        *,
        replay_readings: str | Path | None = None,
    ) -> None:
        self.output_root = Path(output_root).resolve()
        self.config = config or AppConfig()
        self.engine = engine or TesseractOCREngine(
            self.config.tesseract_path, self.config.languages, self.config.timeout_seconds
        )
        self.state_dir = self.output_root / ".ocr_fiscal"
        self.output_root.mkdir(parents=True, exist_ok=True)
        # La salida debe estar vacía: nunca se mezclan resultados de corridas distintas.
        if any(self.output_root.iterdir()):
            raise ValueError("La salida ya contiene archivos. Elija una carpeta vacía o una ruta nueva.")
        # Crear el directorio de estado reserva la salida incluso si dos procesos llegan juntos.
        try:
            self.state_dir.mkdir()
        except FileExistsError as exc:
            raise ValueError("Otro proceso ya reservó la salida. Elija una ruta nueva.") from exc
        self._scan_started = threading.Event()
        self._scan_lock = threading.Lock()
        self.store = StateStore(self.state_dir / "state.sqlite3")
        # Lecturas guardadas: se graban si el ajuste lo pide; con replay_readings se leen de
        # una corrida anterior en lugar de pasar por Tesseract.
        self.replay = ReadingCache(replay_readings, readonly=True) if replay_readings is not None else None
        self.reading_cache = ReadingCache(self.state_dir / CACHE_FILENAME) if self.config.keep_full_reading else None
        self._store_effective_run_metadata()

    # Borra .ocr_fiscal/ (state.sqlite3) cuando el usuario pidió no conservar el estado. Se llama después de
    # escribir el Excel; sin el estado, el comando export ya no puede regenerarlo.
    def discard_state(self) -> None:
        import shutil

        self.close_readings()
        shutil.rmtree(self.state_dir, ignore_errors=False)

    def close_readings(self) -> None:
        for cache in (self.reading_cache, self.replay):
            if cache is not None:
                cache.close()
        self.reading_cache = self.replay = None

    # Guarda los datos de la corrida: RUC del comprador, ajustes de lectura y, al terminar, los contadores finales.
    def _store_effective_run_metadata(self, summary: ScanSummary | None = None) -> None:
        metadata: dict[str, Any] = {
            "run_id": "local-scan",
            "target_recipient_ruc": self.config.recipient_ruc,
            "inspect_archives": self.config.inspect_archives,
            "inspect_nested_archives": self.config.inspect_nested_archives,
            # Guardamos los ajustes de lectura para que dos corridas se puedan comparar
            # sabiendo con qué se hizo cada una.
            "idiomas_ocr": self.config.languages,
            "resolucion_pdf": self.config.render_dpi,
            "umbral_confianza_ocr": round(self.config.min_ocr_confidence * 100),
            "segundos_por_pagina": self.config.timeout_seconds,
            "releer_recuadro_importes": self.config.read_amount_panel,
            "exportar_pdf_ordenado": self.config.export_pdf_folders,
            "padron_ruc": str(self.config.padron_db_path) if self.config.padron_db_path else None,
            "consulta_ruc": str(self.config.consulta_ruc_path) if self.config.consulta_ruc_path else None,
            "guardar_lectura_completa": self.config.keep_full_reading,
            "redetectada_desde": str(self.replay.path) if self.replay is not None else None,
            "input_root": str(getattr(self, "_input_root", "")) or None,
        }
        if summary is not None:
            metadata.update(
                {
                    "run_status": summary.run_status,
                    "cancelled": summary.cancelled,
                    "discovered_unit_count": summary.discovered_unit_count,
                    "started_unit_count": summary.started_unit_count,
                    "processed_file_count": summary.processed_file_count,
                    "pending_unit_count": summary.pending_unit_count,
                    "pending_units": list(summary.pending_units),
                    "excluded_xml_count": summary.excluded_xml_count,
                    "ignored_unsupported_count": summary.ignored_unsupported_count,
                    "padron_summary": dict(summary.padron_summary),
                    "consulta_ruc_summary": dict(summary.consulta_ruc_summary),
                    "reconciliation_summary": dict(summary.reconciliation_summary),
                }
            )
        self.store.store_run_metadata(metadata)

    # Guarda en la base el resultado de un bloque de página y devuelve su DocumentOutcome.
    # También registra los eventos de detección y de identidad.
    def _store_result(
        self,
        unit: WorkUnit,
        result: ExtractionResult,
        page_count: int,
        source_format: str,
        elapsed_seconds: float,
    ) -> DocumentOutcome:
        ref = unit.source_ref

        # En operación normal no retenemos la transcripción completa en SQLite,
        # conservando únicamente los fragmentos de evidencia, ubicaciones y motivos.
        retained_text = ""

        # Se arma el registro con los datos de origen (archivo, ZIP, páginas) y los del resultado.
        record = {
            "source_id": ref.source_id,
            "original_path": ref.source_path,
            "container_chain": list(ref.container_chain),
            "member_path": ref.member_path,
            "nested_depth": ref.nested_depth,
            "display_location": ref.display_location,
            "source_format": source_format,
            "page_count": page_count,
            "page_start": result.page_start,
            "page_end": result.page_end,
            "block_start": result.block_start,
            "block_end": result.block_end,
            **result.as_dict(),
            "text_sidecar_path": None,
            "normalized_filename": None,
            "organized_path": None,
            "extracted_text": retained_text,
        }
        # Se guarda, y se anotan dos eventos: cómo se decidió que era un comprobante y si su identidad quedó completa.
        document_id = self.store.store_document(record)
        scoped_events: list[StageEvent] = []
        scoped_events.append(
            StageEvent(
                stage="receipt_detection",
                status="succeeded",
                message=f"{result.detection_status} ({result.classification_rule})",
                page_number=result.page_start,
                source_id=ref.source_id,
                member_path=ref.member_path,
                container_chain=ref.container_chain,
            )
        )
        scoped_events.append(
            StageEvent(
                stage="identity_validation",
                status="succeeded" if result.identidad_completa else "incomplete",
                message="identidad completa" if result.identidad_completa else "faltan campos o hay candidatos ambiguos",
                page_number=result.page_start,
                source_id=ref.source_id,
                member_path=ref.member_path,
                container_chain=ref.container_chain,
            )
        )
        self.store.add_events(document_id, scoped_events)
        # Se devuelve el resumen del documento ya guardado, copiando los datos del resultado.
        return DocumentOutcome(
            document_id=document_id,
            original_path=ref.source_path,
            detection_status=result.detection_status,
            identity_status=result.identity_status,
            comprobante_detectado=result.comprobante_detectado,
            identidad_completa=result.identidad_completa,
            renombrable=result.renombrable,
            page_start=result.page_start,
            page_end=result.page_end,
            block_start=result.block_start,
            block_end=result.block_end,
            normalized_filename=None,
            output_path=None,
            source_id=ref.source_id,
            container_chain=list(ref.container_chain),
            member_path=ref.member_path,
            nested_depth=ref.nested_depth,
            display_location=ref.display_location,
            reading_source=result.reading_source,
            classification_rule=result.classification_rule,
            non_receipt_type=result.non_receipt_type,
            decision_comment=result.decision_comment,
            detection_reason=result.detection_reason,
            detection_excerpt=result.detection_excerpt,
            identity_reason=result.identity_reason,
            identity_excerpt=result.identity_excerpt,
            recipient_ruc_match=result.recipient_ruc_match,
            recipient_role_status=result.recipient_role_status,
            issuer_role_status=result.issuer_role_status,
            third_party_rucs=list(result.third_party_rucs),
            candidate_fiscal_key=result.candidate_fiscal_key,
            secondary_review_fields=list(result.secondary_review_fields),
            general_description=result.general_description,
            requires_human_review=result.requires_human_review,
            line_items=[it.as_dict() for it in result.line_items],
            derived_checks=list(result.derived_checks),
            field_statuses=dict(result.field_statuses),
            fields=result.fields_normalized,
            fields_original=dict(result.fields_original),
            confidence=dict(result.confidence),
            field_evidence={name: [item.as_dict() for item in values] for name, values in result.field_evidence.items()},
            candidates=dict(result.candidates),
            observations=list(result.observations),
            read_error=result.read_error,
            read_error_details=list(result.read_error_details),
            text_sidecar_path=None,
            elapsed_seconds=elapsed_seconds,
        )

    # Paso final de la corrida, con todos los documentos ya leídos: contrasta los nombres con el padrón (si lo hay)
    # y reúne las copias del mismo comprobante.
    def _reconcile_outcomes(
        self,
        summary: ScanSummary,
        unit_group_counts: Mapping[int, int],
    ) -> None:
        """Contrasta la razón social y conserva una representación por clave concordante."""
        # Paso 1: juntar todos los RUC que aparecieron (emisores, receptores y terceros).
        all_rucs: list[str] = []
        for item in summary.outcomes:
            em = item.fields.get("emitter_ruc") or item.fields_original.get("emitter_ruc")
            rec = item.fields.get("recipient_ruc") or item.fields_original.get("recipient_ruc")
            if em:
                all_rucs.append(str(em))
            if rec:
                all_rucs.append(str(rec))
            for tp in item.third_party_rucs:
                if tp:
                    all_rucs.append(str(tp))

        # Paso 2: buscarlos en el padrón local, si hay uno configurado.
        # Sin padrón, el informe lo dice y no se contrasta nada.
        padron_adapter = LocalPadronAdapter(
            self.config.padron_db_path,
            max_age_days=self.config.padron_max_age_days,
        )
        padron_report = padron_adapter.evaluate_rucs(all_rucs)
        summary.padron_summary = {
            "db_status": padron_report.db_status,
            "db_path": padron_report.db_path,
            "cutoff_date": padron_report.cutoff_date,
            "source_path": padron_report.source_path,
            "queried_count": padron_report.queried_count,
            "present_count": padron_report.present_count,
            "absent_count": padron_report.absent_count,
            "obsolete_count": padron_report.obsolete_count,
            "invalid_ruc_count": padron_report.invalid_ruc_count,
            "reason": padron_report.reason,
        }

        # Paso 2b: con la consulta RUC (si la hay), el CIIU del emisor decide la tasa especial de IGV.
        consulta = load_consulta_ruc(self.config.consulta_ruc_path) if self.config.consulta_ruc_path else None
        if consulta is not None:
            summary.consulta_ruc_summary = {
                "path": consulta.path,
                "query_date": consulta.query_date,
                "ruc_count": len(consulta.activities),
                "rejected_rows": consulta.rejected_rows,
            }
            for item in summary.outcomes:
                _apply_emitter_activity(item, consulta)

        # El padrón local se usa para validar existencia del RUC en la instantánea y cotejar razones sociales.
        for item in summary.outcomes:
            em_ruc = str(item.fields.get("emitter_ruc") or item.fields_original.get("emitter_ruc") or "").strip()
            obs_em_name = item.fields.get("issuer_name") or item.fields_original.get("issuer_name")
            em_match_status = "no_evaluable"
            em_similarity = 0.0
            if em_ruc and em_ruc in padron_report.entries:
                em_entry = padron_report.entries[em_ruc]
                item.emitter_padron_status = em_entry.status
                item.emitter_padron_razon_social = em_entry.razon_social
                item.emitter_padron_estado = em_entry.estado
                item.emitter_padron_cond_domicilio = em_entry.cond_domicilio
                item.emitter_padron_cutoff_date = em_entry.cutoff_date
                item.emitter_padron_source = em_entry.source_path
                if em_entry.status == "ausente" and item.detection_status == DETECTED:
                    if "emitter_ruc_padron" not in item.secondary_review_fields:
                        item.secondary_review_fields.append("emitter_ruc_padron")
                    item.requires_human_review = True
                    item.observations.append(f"ruc_emisor_ausente_en_padron: {em_ruc} no figura en la instantánea local")
                em_match_status, em_similarity = compare_names_with_padron(obs_em_name, em_entry.razon_social, ruc=em_ruc)
                if em_match_status == "discrepancia":
                    item.emitter_name_discrepancy = True
                    if "issuer_name" not in item.secondary_review_fields:
                        item.secondary_review_fields.append("issuer_name")
                    item.observations.append(
                        f"discrepancia_razon_social_emisor: observada={obs_em_name!r} vs padron={em_entry.razon_social!r}"
                    )
                elif em_match_status == "posible_nombre_comercial":
                    item.observations.append(
                        f"posible_nombre_comercial_emisor: observado={obs_em_name!r} vs razon_social_padron={em_entry.razon_social!r}"
                    )

            if self.config.recipient_ruc and em_ruc == self.config.recipient_ruc and item.detection_status == DETECTED:
                if "emitter_ruc_self" not in item.secondary_review_fields:
                    item.secondary_review_fields.append("emitter_ruc_self")
                item.requires_human_review = True
                item.observations.append(
                    f"posible_comprobante_de_venta_propia: el RUC emisor ({em_ruc}) coincide con el RUC receptor objetivo"
                )

            rec_ruc = str(item.fields.get("recipient_ruc") or item.fields_original.get("recipient_ruc") or "").strip()
            obs_rec_name = item.fields.get("recipient_name") or item.fields_original.get("recipient_name")
            rec_match_status = "no_evaluable"
            rec_similarity = 0.0
            if rec_ruc and rec_ruc in padron_report.entries:
                rec_entry = padron_report.entries[rec_ruc]
                item.recipient_padron_status = rec_entry.status
                item.recipient_padron_razon_social = rec_entry.razon_social
                item.recipient_padron_estado = rec_entry.estado
                item.recipient_padron_cond_domicilio = rec_entry.cond_domicilio
                if rec_entry.status == "ausente" and item.detection_status == DETECTED:
                    if "recipient_ruc_padron" not in item.secondary_review_fields:
                        item.secondary_review_fields.append("recipient_ruc_padron")
                    item.requires_human_review = True
                    item.observations.append(f"ruc_receptor_ausente_en_padron: {rec_ruc} no figura en la instantánea local")
                rec_match_status, rec_similarity = compare_names_with_padron(obs_rec_name, rec_entry.razon_social, ruc=rec_ruc)
                if rec_match_status == "discrepancia":
                    if "recipient_name" not in item.secondary_review_fields:
                        item.secondary_review_fields.append("recipient_name")
                    item.observations.append(
                        f"discrepancia_razon_social_receptor: observada={obs_rec_name!r} vs padron={rec_entry.razon_social!r}"
                    )

            if item.detection_status == DETECTED and (em_match_status != "no_evaluable" or rec_match_status != "no_evaluable"):
                item.derived_checks = [
                    chk for chk in item.derived_checks if chk.get("check") != "padron_name_crosscheck"
                ]
                item.derived_checks.append(
                    {
                        "check": "padron_name_crosscheck",
                        "consistent": em_match_status != "discrepancia" and rec_match_status != "discrepancia",
                        "emitter_status": em_match_status,
                        "emitter_similarity": em_similarity,
                        "recipient_status": rec_match_status,
                        "recipient_similarity": rec_similarity,
                    }
                )

            # Verificación oficial: este programa no consulta a SUNAT, por eso todo comprobante detectado queda «pendiente».
            if item.detection_status == DETECTED:
                item.official_validation_status = "pendiente"
                item.official_validation_reason = "Esta aplicación no consulta la validez de emisión."
            else:
                item.official_validation_status = "no_aplica"

        # Paso 7: agrupamos y conciliamos por clave fiscal completa (RUC emisor, tipo, serie y correlativo).
        groups_by_key: dict[str, list[DocumentOutcome]] = {}
        for item in summary.outcomes:
            if (
                item.detection_status == DETECTED
                and item.identity_status == IDENTITY_COMPLETE
                and item.candidate_fiscal_key
            ):
                groups_by_key.setdefault(item.candidate_fiscal_key, []).append(item)
            else:
                item.key_reconciliation_status = (
                    "no_aplica" if item.detection_status != DETECTED else "sin_clave_completa"
                )
                item.key_reconciliation_reason = (
                    "Registro no es un CDP."
                    if item.detection_status != DETECTED
                    else "CDP sin identidad completa confirmada; permanece en cola de revisión."
                )
                item.renombrable = False

        # Contadores para el resumen final.
        reconciled_cdp_keys = 0
        conflicting_cdp_keys = 0
        pending_official_cdp_keys = 0

        # Paso 8: cada grupo son copias del mismo comprobante (misma clave). Si sus datos no concuerdan,
        # todas quedan en revisión; si concuerdan, se elige una como principal.
        for cand_key, group in groups_by_key.items():
            doc_types = {
                str(it.fields.get("document_type") or it.fields_original.get("document_type") or "")
                for it in group
            }
            is_adjustment = bool(doc_types & {"07", "08"})

            # Comprobamos que moneda, fecha, adquirente, referencia y totales coincidan entre copias.
            conflicts: list[str] = []
            currencies = {str(it.fields.get("currency")).strip() for it in group if it.fields.get("currency")}
            if len(currencies) > 1:
                conflicts.append(f"moneda({sorted(currencies)})")
            dates = {str(it.fields.get("issue_date")).strip() for it in group if it.fields.get("issue_date")}
            if len(dates) > 1:
                conflicts.append(f"fecha_emision({sorted(dates)})")
            rec_rucs = {str(it.fields.get("recipient_ruc")).strip() for it in group if it.fields.get("recipient_ruc")}
            if len(rec_rucs) > 1:
                conflicts.append(f"ruc_adquirente({sorted(rec_rucs)})")
            refs = {
                str(it.fields.get("referenced_document")).strip()
                for it in group
                if it.fields.get("referenced_document")
            }
            if len(refs) > 1:
                conflicts.append(f"documento_referencia({sorted(refs)})")

            totals_dec: list[Decimal] = []
            igvs_dec: list[Decimal] = []
            gravs_dec: list[Decimal] = []
            for it in group:
                for fname_chk, target_list in (("total", totals_dec), ("igv", igvs_dec), ("gravado", gravs_dec)):
                    raw_val = it.fields.get(fname_chk)
                    if raw_val:
                        try:
                            target_list.append(Decimal(str(raw_val).replace(",", "")))
                        except InvalidOperation:
                            pass
            if len(totals_dec) > 1 and (max(totals_dec) - min(totals_dec)) > Decimal("0.02"):
                conflicts.append(f"total({sorted(str(d) for d in set(totals_dec))})")
            if len(igvs_dec) > 1 and (max(igvs_dec) - min(igvs_dec)) > Decimal("0.02"):
                conflicts.append(f"igv({sorted(str(d) for d in set(igvs_dec))})")
            if len(gravs_dec) > 1 and (max(gravs_dec) - min(gravs_dec)) > Decimal("0.02"):
                conflicts.append(f"base_gravada({sorted(str(d) for d in set(gravs_dec))})")

            # Si las copias discrepan, ninguna se da por buena: todas quedan en revisión.
            if conflicts:
                if not is_adjustment:
                    conflicting_cdp_keys += 1
                reason_msg = (
                    f"Representaciones en conflicto para la clave {cand_key}: {', '.join(conflicts)}. "
                    "Se retiene el grupo en revisión sin publicar ni copiar."
                )
                for it in group:
                    it.key_reconciliation_status = "en_conflicto"
                    it.key_reconciliation_reason = reason_msg
                    it.identity_status = IDENTITY_CONFLICT
                    it.identidad_completa = False
                    it.renombrable = False
                    it.is_primary_representation = False
                    it.eligible_for_valid_sheet = False
                    it.requires_human_review = True
                    it.observations.append(reason_msg)
                continue

            if not is_adjustment:
                reconciled_cdp_keys += 1

            # Contamos cuántos comprobantes detectados trae cada archivo fuente: una factura con sus propios anexos
            # (un comprobante más guías o constancias) no debe quedar en desventaja frente a una copia suelta en ZIP.
            cdp_count_by_source: dict[int | None, int] = {}
            for out_it in summary.outcomes:
                if out_it.detection_status == DETECTED:
                    cdp_count_by_source[out_it.source_id] = cdp_count_by_source.get(out_it.source_id, 0) + 1

            # Criterio para elegir la copia principal: se prefiere la que venga como único CDP de su archivo,
            # sin error de lectura, con texto digital nativo (aunque venga en ZIP) sobre escaneos,
            # con cuadre aritmético y QR coincidente, con menos campos en revisión, y en disco antes que en ZIP.
            def _primary_sort_key(it: DocumentOutcome) -> tuple[Any, ...]:
                single_cdp_unit = 1 if cdp_count_by_source.get(it.source_id, 1) == 1 else 0
                no_err = 1 if not it.read_error else 0
                is_native = 1 if it.reading_source == "native_pdf_text" else 0
                arith_ok = 1 if any(
                    chk.get("check") == "arithmetic_total_reconciliation" and chk.get("consistent")
                    for chk in it.derived_checks
                ) else 0
                qr_ok = 1 if any(
                    chk.get("check") == "sunat_qr_crosscheck" and chk.get("status") == "coincide"
                    for chk in it.derived_checks
                ) else 0
                few_reviews = len(it.secondary_review_fields)
                on_disk = 1 if (it.nested_depth == 0 and it.member_path is None) else 0
                sec_fields = sum(
                    1
                    for k in ("issue_date", "recipient_ruc", "currency", "igv", "total", "issuer_name")
                    if it.fields.get(k) or it.fields_original.get(k)
                )
                return (
                    -single_cdp_unit,
                    -no_err,
                    -is_native,
                    -arith_ok,
                    -qr_ok,
                    few_reviews,
                    -on_disk,
                    -sec_fields,
                    it.source_id or 0,
                    it.page_start,
                    it.block_start,
                    it.document_id,
                )

            # La primera según ese criterio es la principal; las demás quedan como copias secundarias
            # (no salen en la hoja principal).
            ordered = sorted(group, key=_primary_sort_key)
            primary = ordered[0]

            for it in ordered:
                if it is primary:
                    it.is_primary_representation = True
                    if is_adjustment:
                        it.key_reconciliation_status = (
                            "ajuste_unico" if len(group) == 1 else "ajuste_concordante_primario"
                        )
                        it.key_reconciliation_reason = (
                            "Nota vinculada única."
                            if len(group) == 1
                            else f"Nota vinculada principal entre {len(group)} representaciones concordantes."
                        )
                        it.renombrable = False
                        it.eligible_for_valid_sheet = False
                    else:
                        it.key_reconciliation_status = "unico" if len(group) == 1 else "concordante_primario"
                        it.key_reconciliation_reason = (
                            "Representación única de clave fiscal."
                            if len(group) == 1
                            else f"Representación principal elegida entre {len(group)} copias concordantes."
                        )
                        pending_official_cdp_keys += 1
                        it.renombrable = False
                        it.eligible_for_valid_sheet = False
                else:
                    it.is_primary_representation = False
                    it.key_reconciliation_status = (
                        "ajuste_concordante_secundario" if is_adjustment else "concordante_secundario"
                    )
                    it.key_reconciliation_reason = (
                        f"Copia secundaria concordante de {cand_key} "
                        f"(principal: doc_id={primary.document_id}, origen={primary.display_location})."
                    )
                    it.renombrable = False
                    it.normalized_filename = None
                    it.output_path = None
                    it.eligible_for_valid_sheet = False

        for item in summary.outcomes:
            # Se guarda en la base lo que se calculó en esta etapa (identidad, padrón, verificación y reconciliación).
            self.store.update_document_post_processing(
                item.document_id,
                {
                    "identity_status": item.identity_status,
                    "identidad_completa": item.identidad_completa,
                    "renombrable": item.renombrable,
                    "normalized_filename": item.normalized_filename,
                    "organized_path": item.output_path,
                    "official_validation_status": item.official_validation_status,
                    "official_validation_date": item.official_validation_date,
                    "official_validation_source": item.official_validation_source,
                    "official_validation_reference": item.official_validation_reference,
                    "official_validation_reason": item.official_validation_reason,
                    "official_validation_channel": item.official_validation_channel,
                    "emitter_padron_status": item.emitter_padron_status,
                    "emitter_padron_razon_social": item.emitter_padron_razon_social,
                    "emitter_padron_estado": item.emitter_padron_estado,
                    "emitter_padron_cond_domicilio": item.emitter_padron_cond_domicilio,
                    "emitter_padron_cutoff_date": item.emitter_padron_cutoff_date,
                    "emitter_padron_source": item.emitter_padron_source,
                    "emitter_name_discrepancy": item.emitter_name_discrepancy,
                    "recipient_padron_status": item.recipient_padron_status,
                    "recipient_padron_razon_social": item.recipient_padron_razon_social,
                    "recipient_padron_estado": item.recipient_padron_estado,
                    "recipient_padron_cond_domicilio": item.recipient_padron_cond_domicilio,
                    "key_reconciliation_status": item.key_reconciliation_status,
                    "key_reconciliation_reason": item.key_reconciliation_reason,
                    "is_primary_representation": item.is_primary_representation,
                    "eligible_for_valid_sheet": item.eligible_for_valid_sheet,
                    "requires_human_review": item.requires_human_review,
                    "observations": item.observations,
                    "secondary_review_fields": item.secondary_review_fields,
                    "derived_checks": item.derived_checks,
                },
            )

        # Resumen de la reconciliación. Ningún comprobante se confirma oficialmente (official_confirmed_cdp_keys = 0)
        # porque no hay consulta a SUNAT.
        summary.reconciliation_summary = {
            "official_source_configured": False,
            "reconciled_concordant_cdp_keys": reconciled_cdp_keys,
            "conflicting_cdp_keys": conflicting_cdp_keys,
            "official_confirmed_cdp_keys": 0,
            "pending_official_validation_cdp_keys": pending_official_cdp_keys,
            "organized_primary_copies_count": 0,
        }

    # Resultado para un documento que no se pudo leer: queda como error de lectura y exige revisión humana.
    def _read_failure(self, error: Exception | str) -> ExtractionResult:
        msg = error if isinstance(error, str) else f"{type(error).__name__}: {error}"
        return ExtractionResult(
            detection_status=READ_ERROR,
            identity_status=IDENTITY_NOT_APPLICABLE,
            comprobante_detectado=False,
            identidad_completa=False,
            renombrable=False,
            page_start=1,
            page_end=1,
            classification_rule="read_failure",
            detection_reason=f"Error de lectura documental: {msg}",
            decision_comment=f"Error de lectura documental: {msg}",
            requires_human_review=True,
            read_error=True,
            observations=["no fue posible leer el archivo"],
            read_error_details=[msg],
        )

    def _process_single_unit(
        self,
        unit: WorkUnit,
        target_ruc: str | None,
    ) -> tuple[list[ExtractionResult], list[StageEvent], str, int, float, float, float, dict[int, int]]:
        """Lee, extrae, ejecuta cotejos locales (QR y relectura de línea) para una unidad."""
        unit_started = time.perf_counter()
        read_t0 = time.perf_counter()
        try:
            key: str | None = None
            rereads: dict[str, str | None] = {}
            if self.replay is not None or self.reading_cache is not None:
                content = unit.read_bytes()
                if content is None and unit.disk_path is not None:
                    content = unit.disk_path.read_bytes()
                key = unit_key(content, self.config) if content is not None else None
            if self.replay is not None:
                cached = self.replay.get(key) if key else None
                if cached is None:
                    raise DocumentReadError(
                        "lectura no guardada: el documento no figura en las lecturas de la corrida de origen "
                        "(nuevo, modificado o leído con otros ajustes). Vuelva a correr scan para leerlo."
                    )
                read_result, saved = cached

                def reread_fn(page_number: int, bbox: Any, rotation: int) -> str | None:
                    value = saved.get(reread_key(page_number, bbox))
                    rereads[reread_key(page_number, bbox)] = value
                    return value
            else:
                read_result = read_work_unit(unit, self.engine, self.config)
                reread_fn = None
                if self.reading_cache is not None and isinstance(self.engine, TesseractOCREngine):
                    def reread_fn(page_number: int, bbox: Any, rotation: int) -> str | None:
                        value = reocr_line_bbox(unit, page_number, bbox, self.config, self.engine, rotation_angle=rotation)
                        rereads[reread_key(page_number, bbox)] = value
                        return value
            read_ms = (time.perf_counter() - read_t0) * 1000.0
            rotations = {
                getattr(p, "page_number", 1): int((getattr(p, "quality", {}) or {}).get("rotation_angle", 0))
                for p in read_result.pages
                if int((getattr(p, "quality", {}) or {}).get("rotation_angle", 0)) in {90, 180, 270}
            }

            ext_t0 = time.perf_counter()
            results = parse_document(
                read_result.pages,
                self.config.document_types,
                self.config.min_ocr_confidence,
                target_recipient_ruc=target_ruc,
                rounding_tolerance=self.config.rounding_tolerance,
            )
            events = list(read_result.events)

            if (
                self.replay is None
                and self.config.read_amount_panel and read_result.source_format == "pdf"
                and len(results) == 1
                and results[0].comprobante_detectado and results[0].document_type in {"01", "03", "07", "08"}
            ):
                total_ev = results[0].field_evidence.get("total") or []
                target_page_no = total_ev[0].page_number if total_ev else results[0].page_end
                page_map = {p.page_number: p for p in read_result.pages}
                target_page = page_map.get(target_page_no) or (read_result.pages[-1] if read_result.pages else None)
                if target_page is not None:
                    panel_event = review_amount_panel(unit, target_page, results[0], self.config, self.engine)
                    if panel_event is not None:
                        events.append(panel_event)

            events.extend(_surgical_line_reocr(unit, read_result.pages, results, self.config, self.engine, reread_fn))
            if self.reading_cache is not None and key:
                self.reading_cache.put(key, read_result, rereads)
            for item_res in results:
                _crosscheck_sunat_qr(item_res, read_result.pages)

            ext_ms = (time.perf_counter() - ext_t0) * 1000.0
            return results, events, read_result.source_format, read_result.page_count, read_ms, ext_ms, (time.perf_counter() - unit_started), rotations
        except Exception as exc:
            read_ms = (time.perf_counter() - read_t0) * 1000.0
            result = self._read_failure(exc)
            events = [
                StageEvent(
                    "document_read",
                    "failed",
                    f"{type(exc).__name__}: {exc}",
                    1,
                    source_id=unit.source_ref.source_id,
                    member_path=unit.source_ref.member_path,
                    container_chain=unit.source_ref.container_chain,
                )
            ]
            return [result], events, unit.source_ref.detected_format, 0, read_ms, 0.0, (time.perf_counter() - unit_started), {}

    # La corrida completa: recorre la entrada, lee cada documento, extrae sus datos, los guarda y, al final, los reconcilia.
    # Se puede cancelar y va avisando de su avance.
    def scan(
        self,
        input_root: str | Path,
        *,
        progress_callback: Callable[[ProgressUpdate], None] | None = None,
        cancel_event: threading.Event | None = None,
    ) -> ScanSummary:
        # Paso 1: comprobar la entrada y reservar la salida (una salida sólo admite una corrida).
        root = Path(input_root).resolve()
        if not root.is_dir():
            raise ValueError(f"La carpeta de entrada no existe: {root}")
        self._input_root = root

        if root == self.output_root:
            raise ValueError("La carpeta de salida debe ser distinta de la carpeta de entrada.")
        with self._scan_lock:
            if self._scan_started.is_set():
                raise ValueError("Esta salida ya tiene una corrida. Elija una ruta nueva.")
            self._scan_started.set()

        # Paso 2: cronómetros por etapa, para saber en qué se va el tiempo.
        stage_timings_ms: dict[str, float] = {
            "discovery_and_archives": 0.0,
            "document_read_and_ocr": 0.0,
            "extraction_and_triage": 0.0,
            "storage": 0.0,
        }

        # Paso 3: recorrer la carpeta y los ZIP, y guardar el inventario de todo lo encontrado.
        t0 = time.perf_counter()
        walker = ContainerWalker(self.config, output_root=self.output_root)
        walk_result = walker.walk(root, recursive=True)
        stage_timings_ms["discovery_and_archives"] = (time.perf_counter() - t0) * 1000.0

        self.store.store_inventory(walk_result.inventory)
        self.store.add_events(None, walk_result.events)

        # Paso 4: preparar el resumen de la corrida con los primeros contadores.
        total_units = len(walk_result.work_units)
        zip_members = sum(1 for inv in walk_result.inventory if inv.nested_depth > 0 and inv.detected_format != "directory")
        summary = ScanSummary(
            discovered_unit_count=total_units,
            processed_file_count=0,
            pending_unit_count=0,
            excluded_xml_count=walk_result.excluded_xml_count,
            ignored_unsupported_count=walk_result.ignored_unsupported_count,
            top_level_file_count=walk_result.top_level_file_count,
            zip_container_count=len(walk_result.archive_summaries),
            zip_member_count=zip_members,
            archive_summaries=walk_result.archive_summaries,
        )

        # Para poder atribuir a cada ZIP lo que se encontró dentro de él.
        archive_by_key: dict[tuple[str, tuple[str, ...]], ArchiveSummary] = {
            (a.source_path, tuple(a.container_chain)): a for a in walk_result.archive_summaries
        }
        processed_unit_indices: set[int] = set()
        unit_group_counts: dict[int, int] = {}
        units_by_source_id: dict[int, WorkUnit] = {}
        source_formats_by_doc: dict[int, str] = {}
        rotations_by_source_id: dict[int, dict[int, int]] = {}
        target_ruc = self.config.recipient_ruc

        def _persist_unit_output(idx: int, unit: WorkUnit, payload: tuple[list[ExtractionResult], list[StageEvent], str, int, float, float, float, dict[int, int]]) -> None:
            results, events, source_format, page_count, read_ms, ext_ms, elapsed, rotations = payload
            stage_timings_ms["document_read_and_ocr"] += read_ms
            stage_timings_ms["extraction_and_triage"] += ext_ms
            per_result_elapsed = elapsed / max(1, len(results))
            store_t0 = time.perf_counter()
            ref = unit.source_ref
            units_by_source_id[ref.source_id] = unit
            if rotations:
                rotations_by_source_id[ref.source_id] = rotations
            self.store.add_events(None, [
                replace(event, source_id=ref.source_id, member_path=ref.member_path,
                        container_chain=ref.container_chain)
                for event in events
            ])
            for result in results:
                outcome = self._store_result(unit, result, page_count, source_format, per_result_elapsed)
                unit_group_counts[outcome.document_id] = len(results)
                source_formats_by_doc[outcome.document_id] = source_format
                summary.outcomes.append(outcome)

                if unit.source_ref.container_chain:
                    for chain_len in range(1, len(unit.source_ref.container_chain) + 1):
                        sub_chain = tuple(unit.source_ref.container_chain[:chain_len])
                        arch = archive_by_key.get((unit.source_ref.source_path, sub_chain))
                        if arch is not None:
                            if outcome.detection_status == READ_ERROR or outcome.read_error:
                                arch.read_errors += 1
                            elif outcome.detection_status == DETECTED:
                                arch.detected_receipts += 1
                            elif outcome.detection_status == AMBIGUOUS:
                                arch.ambiguous_records += 1
                            elif outcome.detection_status == NOT_RECEIPT:
                                arch.non_receipts += 1

            summary.processed_file_count += 1
            processed_unit_indices.add(idx - 1)
            if unit.source_ref.container_chain:
                for chain_len in range(1, len(unit.source_ref.container_chain) + 1):
                    sub_chain = tuple(unit.source_ref.container_chain[:chain_len])
                    arch = archive_by_key.get((unit.source_ref.source_path, sub_chain))
                    if arch is not None:
                        arch.processed_members += 1
            stage_timings_ms["storage"] += (time.perf_counter() - store_t0) * 1000.0

            if progress_callback is not None:
                progress_callback(
                    ProgressUpdate(
                        stage="procesando",
                        current_unit=summary.processed_file_count,
                        total_units=total_units,
                        current_path=unit.display_path,
                        detected_count=summary.detected_count,
                        ambiguous_count=summary.ambiguous_count,
                        no_comprobante_count=summary.no_comprobante_count,
                        read_error_count=summary.read_error_count,
                        zip_container_count=summary.zip_container_count,
                        message=f"Procesado {Path(unit.source_ref.source_path).name}",
                    )
                )

        # Paso 5: lectura y extracción (en paralelo en CPU cuando hay múltiples unidades; guardado en orden determinista).
        max_workers = min(max(1, (os.cpu_count() or 2) - 1), 4)
        work_units = walk_result.work_units
        if max_workers <= 1 or len(work_units) <= 1:
            for idx, unit in enumerate(work_units, start=1):
                if cancel_event is not None and cancel_event.is_set():
                    summary.cancelled = True
                    summary.run_status = "cancelado"
                    self.store.add_events(
                        None,
                        [
                            StageEvent(
                                "scan_cancelled",
                                "cancelled",
                                f"Cancelado por el usuario tras procesar {summary.processed_file_count}/{total_units} unidades",
                            )
                        ],
                    )
                    break
                summary.started_unit_count += 1
                _persist_unit_output(idx, unit, self._process_single_unit(unit, target_ruc))
        else:
            with ThreadPoolExecutor(max_workers=max_workers) as pool:
                in_flight: dict[Any, tuple[int, WorkUnit]] = {}
                completed_buffer: dict[int, tuple[WorkUnit, tuple[list[ExtractionResult], list[StageEvent], str, int, float, float, float]]] = {}
                next_submit = 0
                next_commit = 1

                while next_submit < len(work_units) and len(in_flight) < max_workers * 2:
                    if cancel_event is not None and cancel_event.is_set():
                        break
                    summary.started_unit_count += 1
                    u_cur = work_units[next_submit]
                    fut = pool.submit(self._process_single_unit, u_cur, target_ruc)
                    in_flight[fut] = (next_submit + 1, u_cur)
                    next_submit += 1

                while in_flight:
                    if cancel_event is not None and cancel_event.is_set():
                        summary.cancelled = True
                        summary.run_status = "cancelado"
                        self.store.add_events(
                            None,
                            [
                                StageEvent(
                                    "scan_cancelled",
                                    "cancelled",
                                    f"Cancelado por el usuario tras procesar {summary.processed_file_count}/{total_units} unidades",
                                )
                            ],
                        )
                        for fut in list(in_flight):
                            fut.cancel()
                        break

                    done, _ = wait(in_flight.keys(), timeout=0.25, return_when=FIRST_COMPLETED)
                    for fut in done:
                        idx_done, u_done = in_flight.pop(fut)
                        completed_buffer[idx_done] = (u_done, fut.result())
                        while next_submit < len(work_units) and len(in_flight) < max_workers * 2:
                            if cancel_event is not None and cancel_event.is_set():
                                break
                            summary.started_unit_count += 1
                            u_next = work_units[next_submit]
                            new_fut = pool.submit(self._process_single_unit, u_next, target_ruc)
                            in_flight[new_fut] = (next_submit + 1, u_next)
                            next_submit += 1

                    while next_commit in completed_buffer:
                        u_commit, payload_commit = completed_buffer.pop(next_commit)
                        _persist_unit_output(next_commit, u_commit, payload_commit)
                        next_commit += 1

        # Registrar unidades pendientes y distinguir contenedor inspeccionado vs miembros procesados
        unprocessed_units = [
            u for u_idx, u in enumerate(walk_result.work_units) if u_idx not in processed_unit_indices
        ]
        summary.pending_unit_count = len(unprocessed_units)
        summary.pending_units = [u.display_path for u in unprocessed_units]
        if summary.cancelled:
            summary.run_status = "cancelado"
        else:
            summary.run_status = "completado"

        for u in unprocessed_units:
            if u.source_ref.container_chain:
                member_label = u.source_ref.member_path or " -> ".join(u.source_ref.container_chain)
                for chain_len in range(1, len(u.source_ref.container_chain) + 1):
                    sub_chain = tuple(u.source_ref.container_chain[:chain_len])
                    arch = archive_by_key.get((u.source_ref.source_path, sub_chain))
                    if arch is not None:
                        if member_label not in arch.pending_members:
                            arch.pending_members.append(member_label)

        for arch in walk_result.archive_summaries:
            arch.unprocessed_members_count = len(arch.pending_members)
            if summary.cancelled and arch.unprocessed_members_count > 0:
                arch.cancelled = True

        # Paso 6: reconciliar (padrón y copias del mismo comprobante), organizar PDFs por proveedor (si está activado),
        # guardar el resumen de los ZIP y los datos finales de la corrida.
        reconcile_t0 = time.perf_counter()
        self._reconcile_outcomes(summary, unit_group_counts)
        if self.config.export_pdf_folders:
            pdf_stats = organize_pdf_expediente(
                self.output_root,
                summary.outcomes,
                units_by_source_id,
                source_formats_by_doc,
                rotations_by_source_id,
                unit_group_counts,
                max_provider_name_len=self.config.max_provider_folder_name_len,
                unify_notes_by_provider=self.config.unify_notes_by_provider,
                target_recipient_ruc=self.config.recipient_ruc,
            )
            for item in summary.outcomes:
                self.store.update_document_post_processing(
                    item.document_id,
                    {
                        "normalized_filename": item.normalized_filename,
                        "organized_path": item.output_path,
                        "observations": item.observations,
                        "requires_human_review": item.requires_human_review,
                        "derived_checks": item.derived_checks,
                        "secondary_review_fields": item.secondary_review_fields,
                    },
                )
            summary.reconciliation_summary["organized_primary_copies_count"] = (
                pdf_stats["cdp_complete"] + pdf_stats["notes_complete"]
            )
            summary.reconciliation_summary["pdf_export_summary"] = pdf_stats
            self.store.add_events(
                None,
                [
                    StageEvent(
                        "pdf_expediente_export",
                        "succeeded",
                        f"Exportados {pdf_stats['total_exported']} PDFs (cdp={pdf_stats['cdp_complete']}, notas={pdf_stats['notes_complete']}, revision={pdf_stats['pending_review']})",
                    )
                ],
            )
        stage_timings_ms["reconciliation_and_validation"] = (time.perf_counter() - reconcile_t0) * 1000.0

        self.store.store_archive_summaries(walk_result.archive_summaries)
        if walk_result.spool_dir is not None:
            import shutil

            shutil.rmtree(walk_result.spool_dir, ignore_errors=True)
        summary.stage_timings_ms = stage_timings_ms
        self._store_effective_run_metadata(summary)

        # Último aviso de avance: completado o cancelado.
        if progress_callback is not None:
            progress_callback(
                ProgressUpdate(
                    stage="cancelado" if summary.cancelled else "completado",
                    current_unit=summary.processed_file_count,
                    total_units=total_units,
                    current_path=str(root),
                    detected_count=summary.detected_count,
                    ambiguous_count=summary.ambiguous_count,
                    no_comprobante_count=summary.no_comprobante_count,
                    read_error_count=summary.read_error_count,
                    zip_container_count=summary.zip_container_count,
                    message=(
                        f"Ejecución cancelada: {summary.processed_file_count}/{total_units} unidades procesadas, "
                        f"{summary.pending_unit_count} pendientes"
                        if summary.cancelled
                        else f"Procesamiento finalizado ({summary.processed_file_count}/{total_units} unidades)"
                    ),
                )
            )
        return summary
