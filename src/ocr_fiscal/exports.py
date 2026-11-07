from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import json
import re
import unicodedata
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Mapping, Sequence

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from .amounts import AMOUNT_COLUMNS, SETTLEMENT_FIELDS
from .extraction import document_title_from_line, has_document_title
from .models import (
    AMBIGUOUS,
    DETECTED,
    IDENTITY_COMPLETE,
    IDENTITY_CONFLICT,
    IDENTITY_INCOMPLETE,
    IDENTITY_NOT_APPLICABLE,
    NOT_RECEIPT,
    READ_ERROR,
    normalize_detection_status,
)
from .processor import ScanSummary
from .storage import StateStore
from .tax import normalize_correlativo
from .visible import visible_labels


# Nombre del libro de Excel que se genera.
OUTPUT_FILENAME = "CDP.xlsx"
# Las fechas se guardan como fechas reales de Excel y se muestran como día/mes/año.
DATE_FORMAT = "dd/mm/yyyy"
# Columnas que contienen fechas y que se convierten en fechas de Excel.
DATE_KEYS = {"valid_issue_date", "issue_date_normalized", "official_validation_date", "emitter_padron_cutoff_date"}
# Nombre de la columna de moneda que acompaña a la detracción, la retención y la percepción.
SETTLEMENT_CURRENCY_EXPORT = {
    "importe_detraccion": "Moneda de detracción",
    "importe_retencion": "Moneda de retención",
    "importe_percepcion": "Moneda de percepción",
}
# Columnas de importes que se exportan: los importes leídos, el total recalculado, la diferencia,
# el cuadre y de dónde salió ese cuadre.
AMOUNTS_EXPORT = [
    *[(label, f"importe_{field}") for label, field in AMOUNT_COLUMNS if field not in SETTLEMENT_FIELDS],
    ("Total calculado", "importe_calculado"),
    ("Diferencia del total", "importe_diferencia"),
    ("Cuadre de importes", "estado_importes"),
    ("Fuente del cuadre", "fuente_importes"),
]
# Detracción, retención y percepción: cada importe va seguido de su propia moneda.
SETTLEMENT_EXPORT = [
    column
    for label, field in AMOUNT_COLUMNS
    if field in SETTLEMENT_FIELDS
    for column in ((label, f"importe_{field}"), (SETTLEMENT_CURRENCY_EXPORT[field], f"moneda_{field}"))
]


# Columnas de la hoja principal (cdp_detectados): el título que lleva en el Excel y el nombre interno que la alimenta.
# El nombre «CDP válidos» es heredado: la hoja sólo lista comprobantes detectados y nada aquí acredita su validez oficial.
CDP_VALIDOS_COLUMNS: list[tuple[str, str]] = [
    ("RUC emisor", "valid_emitter_ruc"),
    ("Razón social observada", "valid_issuer_name_observed"),
    ("Tipo", "valid_document_type"),
    ("Denominación observada", "document_title"),
    ("Serie", "valid_series"),
    ("Correlativo", "valid_correlativo"),
    ("Fecha de emisión", "valid_issue_date"),
    ("RUC adquirente", "valid_recipient_ruc"),
    ("Moneda", "valid_currency"),
    *AMOUNTS_EXPORT,
    *SETTLEMENT_EXPORT,
    ("Estado oficial", "official_validation_status"),
    ("Fecha de verificación", "official_validation_date"),
    ("Archivo de origen", "valid_source_locator"),
    ("Página", "valid_page_locator"),
    ("Descripción observada", "valid_description_observed"),
    ("Detracción visible", "detraccion_visible"),
    ("Retención visible", "retencion_visible"),
]


# Columnas de la hoja notas (notas de crédito y de débito).
AJUSTES_COLUMNS: list[tuple[str, str]] = [
    ("ID documento", "document_id"),
    ("Ubicación completa", "display_location"),
    ("Archivo origen", "original_path"),
    ("Cadena contenedor", "container_chain_str"),
    ("Ruta miembro ZIP", "member_path"),
    ("Páginas", "page_range"),
    ("Bloques", "block_range"),
    ("Fuente de lectura", "reading_source"),
    ("Estado detección", "detection_status"),
    ("Estado identidad", "identity_status"),
    ("Clave fiscal candidata", "candidate_fiscal_key"),
    ("Tipo nota", "document_type_normalized"),
    ("Denominación observada", "document_title"),
    ("Evidencia de la denominación", "document_title_evidence"),
    ("Descripción tipo nota", "note_type_label"),
    ("Serie nota", "series_normalized"),
    ("Correlativo nota", "correlativo_normalized"),
    ("Fecha emisión", "issue_date_normalized"),
    ("RUC emisor", "emitter_ruc_normalized"),
    ("Razón social emisor", "issuer_name_normalized"),
    ("RUC adquirente", "recipient_ruc_normalized"),
    ("Rol receptor", "recipient_role_status"),
    ("Coincide RUC objetivo", "recipient_ruc_match"),
    ("Comprobante afectado leído", "referenced_document_normalized"),
    ("Página comprobante afectado", "referenced_document_page"),
    ("Fragmento comprobante afectado", "referenced_document_excerpt"),
    ("Comprobante afectado en lote", "affected_present_in_batch"),
    ("Ubicaciones comprobante afectado", "affected_locations_in_batch"),
    ("Moneda", "currency_normalized"),
    *AMOUNTS_EXPORT,
    *SETTLEMENT_EXPORT,
    ("Glosa / Motivo ajuste", "general_description"),
    ("Observaciones", "observations"),
]

# Columnas de la hoja revision.
REVIEW_COLUMNS: list[tuple[str, str]] = [
    ("ID documento", "document_id"),
    ("Tipo hallazgo", "finding_type"),
    ("Prioridad", "priority"),
    ("Estado detección", "detection_status"),
    ("Estado identidad", "identity_status"),
    ("Campo a revisar", "review_field"),
    ("Valor observado", "observed_value"),
    ("Página verificación", "verification_page"),
    ("Fragmento evidencia", "evidence_excerpt"),
    ("Motivo de revisión / Comentario", "review_reason"),
    ("Ubicación completa", "display_location"),
    ("Archivo origen", "original_path"),
    ("Cadena contenedor", "container_chain_str"),
    ("Ruta miembro ZIP", "member_path"),
    ("Páginas", "page_range"),
    ("Clave fiscal candidata", "candidate_fiscal_key"),
    ("RUC emisor", "emitter_ruc"),
    ("Rol emisor", "issuer_role_status"),
    ("RUC adquirente", "recipient_ruc"),
    ("Rol receptor", "recipient_role_status"),
    ("Coincide RUC objetivo", "recipient_ruc_match"),
    ("Tipo-Serie-Correlativo", "document_identity"),
    ("Total normalizado", "total_normalized"),
    ("Candidatos / Evidencia", "candidates_or_evidence"),
]

# Columnas de la hoja inventario.
INVENTORY_COLUMNS: list[tuple[str, str]] = [
    ("ID origen", "source_id"),
    ("Tipo registro", "record_kind"),
    ("Archivo origen en disco", "source_path"),
    ("Cadena contenedor", "container_chain_str"),
    ("Ruta miembro ZIP", "member_path"),
    ("Profundidad anidamiento", "nested_depth"),
    ("Ubicación completa", "display_location"),
    ("Páginas / Bloques", "page_block_range"),
    ("Formato detectado", "detected_format"),
    ("Estado inventario", "inventory_status"),
    ("Estado detección", "detection_status"),
    ("Clasificación soporte / Exclusión", "exclusion_or_support_type"),
    ("Motivo / Detalle", "status_reason"),
    ("Acción tomada", "action_taken"),
    ("Tamaño bytes", "byte_size"),
    ("Tamaño comprimido bytes", "compressed_size"),
]

# Convierte a JSON ordenado y con tildes, para guardar listas y diccionarios dentro de una celda.
def _json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


# La clave fiscal de un comprobante: emisor, tipo, serie y correlativo. Sólo existe si su identidad está completa.
def _compute_candidate_key(doc: Mapping[str, Any]) -> str | None:
    """Usa únicamente la clave fiscal emitida por el validador de extracción cuando la identidad es completa."""
    if doc.get("identity_status") != IDENTITY_COMPLETE:
        return None
    cand = doc.get("candidate_fiscal_key")
    if cand and str(cand).strip():
        return str(cand).strip()
    return None


# Pasa un importe de texto a número para que Excel lo trate como cifra.
# Si está vacío o no se entiende, devuelve None: la celda queda vacía, nunca en cero.
def _numeric_or_none(value: Any) -> float | None:
    if value is None or str(value).strip() == "":
        return None
    try:
        return float(Decimal(str(value).replace(",", "").strip()))
    except (InvalidOperation, ValueError):
        return None


# Si se exporta una corrida antigua que no tenía la reconciliación de copias, aquí se rehace en memoria
# con las mismas reglas del procesador: agrupar por clave, detectar conflictos y elegir una copia principal.
def _ensure_reconciled_documents(documents: list[dict[str, Any]]) -> None:
    """Garantiza conciliación determinista en memoria si se exporta desde registros sin post-procesamiento."""
    # Sólo hace falta si hay comprobantes con clave completa a los que todavía no se les calculó la reconciliación.
    needs_reconciliation = any(
        normalize_detection_status(doc.get("detection_status")) == DETECTED
        and doc.get("identity_status") == IDENTITY_COMPLETE
        and doc.get("candidate_fiscal_key")
        and doc.get("key_reconciliation_status") in {None, "", "no_aplica"}
        for doc in documents
    )
    if not needs_reconciliation:
        return

    # Cuántos bloques salieron de cada página física (una página puede traer varios).
    unit_counts: dict[Any, int] = {}
    for doc in documents:
        key_loc = doc.get("source_id") or doc.get("display_location") or doc.get("original_path")
        unit_counts[key_loc] = unit_counts.get(key_loc, 0) + 1

    # Se agrupan las copias del mismo comprobante (misma clave fiscal).
    groups_by_key: dict[str, list[dict[str, Any]]] = {}
    for doc in documents:
        det_st = normalize_detection_status(doc.get("detection_status"))
        cand = _compute_candidate_key(doc)
        if det_st == DETECTED and cand:
            groups_by_key.setdefault(cand, []).append(doc)
        else:
            doc.setdefault(
                "key_reconciliation_status",
                "no_aplica" if det_st != DETECTED else "sin_clave_completa",
            )
            doc.setdefault(
                "key_reconciliation_reason",
                "Registro no es un CDP."
                if det_st != DETECTED
                else "CDP sin identidad completa confirmada; permanece en cola de revisión.",
            )
            doc.setdefault("is_primary_representation", False)
            doc.setdefault("eligible_for_valid_sheet", False)

    # Para cada grupo: si los datos de sus copias no concuerdan, todas quedan en conflicto;
    # si concuerdan, se elige una como principal.
    for cand_key, group in groups_by_key.items():
        doc_types = {
            str((it.get("fields_normalized") or {}).get("document_type") or (it.get("fields_original") or {}).get("document_type") or "")
            for it in group
        }
        is_adjustment = bool(doc_types & {"07", "08"})
        conflicts: list[str] = []
        currencies = {
            str((it.get("fields_normalized") or {}).get("currency")).strip()
            for it in group
            if (it.get("fields_normalized") or {}).get("currency")
        }
        if len(currencies) > 1:
            conflicts.append(f"moneda({sorted(currencies)})")
        dates = {
            str((it.get("fields_normalized") or {}).get("issue_date")).strip()
            for it in group
            if (it.get("fields_normalized") or {}).get("issue_date")
        }
        if len(dates) > 1:
            conflicts.append(f"fecha_emision({sorted(dates)})")
        rec_rucs = {
            str((it.get("fields_normalized") or {}).get("recipient_ruc")).strip()
            for it in group
            if (it.get("fields_normalized") or {}).get("recipient_ruc")
        }
        if len(rec_rucs) > 1:
            conflicts.append(f"ruc_adquirente({sorted(rec_rucs)})")
        totals_dec: list[Decimal] = []
        for it in group:
            raw_tot = (it.get("fields_normalized") or {}).get("total")
            if raw_tot:
                try:
                    totals_dec.append(Decimal(str(raw_tot).replace(",", "")))
                except InvalidOperation:
                    pass
        if len(totals_dec) > 1 and (max(totals_dec) - min(totals_dec)) > Decimal("0.02"):
            conflicts.append(f"total({sorted(str(d) for d in set(totals_dec))})")

        if conflicts:
            reason_msg = (
                f"Representaciones en conflicto para la clave {cand_key}: {', '.join(conflicts)}. "
                "Se retiene el grupo en revisión sin publicar ni copiar."
            )
            for it in group:
                it["key_reconciliation_status"] = "en_conflicto"
                it["key_reconciliation_reason"] = reason_msg
                it["identity_status"] = IDENTITY_CONFLICT
                it["identidad_completa"] = False
                it["renombrable"] = False
                it["is_primary_representation"] = False
                it["eligible_for_valid_sheet"] = False
                it["requires_human_review"] = True
            continue

        # El mismo criterio que usa el procesador para elegir la copia principal.
        def _sort_key(it: Mapping[str, Any]) -> tuple[Any, ...]:
            key_loc = it.get("source_id") or it.get("display_location") or it.get("original_path")
            single_unit = 1 if unit_counts.get(key_loc, 1) == 1 else 0
            no_err = 1 if not it.get("read_error") else 0
            on_disk = 1 if (int(it.get("nested_depth") or 0) == 0 and it.get("member_path") is None) else 0
            is_native = 1 if it.get("reading_source") == "native_pdf_text" else 0
            norm = it.get("fields_normalized") or {}
            orig = it.get("fields_original") or {}
            sec_fields = sum(
                1
                for k in ("issue_date", "recipient_ruc", "currency", "igv", "total", "issuer_name")
                if norm.get(k) or orig.get(k)
            )
            return (
                -single_unit,
                -no_err,
                -on_disk,
                -is_native,
                -sec_fields,
                int(it.get("source_id") or 0),
                int(it.get("page_start") or 1),
                int(it.get("block_start") or 1),
                int(it.get("id") or 0),
            )

        ordered = sorted(group, key=_sort_key)
        primary = ordered[0]
        for it in ordered:
            if it is primary:
                it["is_primary_representation"] = True
                if is_adjustment:
                    it["key_reconciliation_status"] = (
                        "ajuste_unico" if len(group) == 1 else "ajuste_concordante_primario"
                    )
                    it["key_reconciliation_reason"] = (
                        "Nota vinculada única."
                        if len(group) == 1
                        else f"Nota vinculada principal entre {len(group)} representaciones concordantes."
                    )
                    it["eligible_for_valid_sheet"] = False
                else:
                    it["key_reconciliation_status"] = "unico" if len(group) == 1 else "concordante_primario"
                    it["key_reconciliation_reason"] = (
                        "Representación única de clave fiscal."
                        if len(group) == 1
                        else f"Representación principal elegida entre {len(group)} copias concordantes."
                    )
                    it["eligible_for_valid_sheet"] = False
            else:
                it["is_primary_representation"] = False
                it["key_reconciliation_status"] = (
                    "ajuste_concordante_secundario" if is_adjustment else "concordante_secundario"
                )
                it["key_reconciliation_reason"] = (
                    f"Copia secundaria concordante de {cand_key} "
                    f"(principal: doc_id={primary.get('id')}, origen={primary.get('display_location')})."
                )
                it["eligible_for_valid_sheet"] = False


# Arma, para cada clave fiscal, un resumen: cuántas copias hay y en qué archivos y páginas están.
def _build_candidate_key_groups(documents: Sequence[Mapping[str, Any]]) -> dict[str, dict[str, Any]]:
    groups: dict[str, list[Mapping[str, Any]]] = {}
    for doc in documents:
        if normalize_detection_status(doc.get("detection_status")) != DETECTED:
            continue
        key = _compute_candidate_key(doc)
        if not key:
            continue
        groups.setdefault(key, []).append(doc)

    summary_by_key: dict[str, dict[str, Any]] = {}
    for idx, (key, members) in enumerate(sorted(groups.items()), start=1):
        locations: list[str] = []
        primary_doc_id = None
        reconciliation_st = None
        official_st = None
        for m in members:
            loc = f"{m.get('display_location') or m.get('original_path')} (pág. {m.get('page_start', 1)}-{m.get('page_end', 1)})"
            if loc not in locations:
                locations.append(loc)
            if m.get("is_primary_representation"):
                primary_doc_id = m.get("id")
                reconciliation_st = m.get("key_reconciliation_status")
                official_st = m.get("official_validation_status")
        if reconciliation_st is None and members:
            reconciliation_st = members[0].get("key_reconciliation_status")
            official_st = members[0].get("official_validation_status")
        summary_by_key[key] = {
            "group_id": f"CLAVE-{idx:04d}",
            "candidate_fiscal_key": key,
            "count": len(members),
            "locations": locations,
            "locations_str": " | ".join(locations),
            "label": "misma_clave_candidata" if len(members) > 1 else "clave_unica",
            "primary_document_id": primary_doc_id,
            "reconciliation_status": reconciliation_st,
            "official_validation_status": official_st,
        }
    return summary_by_key


# De la evidencia de un campo devuelve la página (y la línea) y el fragmento de la primera lectura.
# Es lo que se muestra en la hoja de revisión para ir directo al original.
def _first_evidence_info(evidence_map: Mapping[str, Any], field_name: str, default_page: str) -> tuple[str, str]:
    items = evidence_map.get(field_name) or []
    if not items:
        return default_page, ""
    first = items[0] if isinstance(items[0], Mapping) else {}
    page_num = first.get("page_number")
    line_num = first.get("line_number")
    page_str = f"pág. {page_num}" + (f", lín. {line_num}" if line_num else "") if page_num else default_page
    excerpt = str(first.get("excerpt") or "")
    return page_str, excerpt


def _document_title_fields(evidence: Mapping[str, Any], document_type: Any) -> dict[str, Any]:
    """Denominación observada y la línea completa que la respalda.

    En bases anteriores el valor guardado era la línea entera, así que recortamos también
    al exportar: el resultado es el mismo título observado, sin añadir ni suponer nada.
    """
    items = evidence.get("document_type") or []
    first = items[0] if items and isinstance(items[0], Mapping) else {}
    line = str(first.get("excerpt") or "")
    observed = str(first.get("observed_value") or line)
    code = str(document_type) if document_type else None
    title = document_title_from_line(observed, code)
    return {
        "document_title": title or None,
        "document_title_evidence": (line.strip() or None) if line.strip() != title else None,
    }


# Texto corto con la página (y el bloque, si hay más de uno) donde está el comprobante.
def _page_locator_str(document: Mapping[str, Any]) -> str:
    p_start = int(document.get("page_start") or 1)
    p_end = int(document.get("page_end") or p_start)
    b_start = int(document.get("block_start") or 1)
    b_end = int(document.get("block_end") or b_start)
    page_part = str(p_start) if p_start == p_end else f"{p_start}-{p_end}"
    if b_start > 1 or b_end > 1:
        block_part = str(b_start) if b_start == b_end else f"{b_start}-{b_end}"
        return f"{page_part} (bloque {block_part})"
    return page_part


# Convierte un documento guardado en una fila plana con todo lo que cualquier hoja pueda necesitar.
# Después cada hoja elige las columnas que le interesan.
def _row(
    document: dict[str, Any],
    key_groups: Mapping[str, Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    original = document.get("fields_original") or {}
    normalized = document.get("fields_normalized") or {}
    evidence = document.get("field_evidence") or {}
    chain = document.get("container_chain") or []
    line_items = document.get("line_items") or []
    # Resumen de los primeros cinco ítems del detalle.
    summary_parts = [
        f"{it.get('description_normalized') or it.get('description_observed') or ''} ({it.get('line_total_normalized') or it.get('line_total_observed') or 's/m'})"
        for it in line_items[:5]
    ]
    if len(line_items) > 5:
        summary_parts.append(f"... (+{len(line_items) - 5} ítems)")

    # Estados y clave fiscal del documento.
    det_status = normalize_detection_status(document.get("detection_status"))
    is_cdp = bool(document.get("comprobante_detectado") or det_status == DETECTED)
    id_completa = bool(document.get("identidad_completa"))
    id_status = str(
        document.get("identity_status")
        or ((IDENTITY_COMPLETE if id_completa else IDENTITY_INCOMPLETE) if is_cdp else IDENTITY_NOT_APPLICABLE)
    )
    cand_key = _compute_candidate_key(document) if is_cdp else None
    grp = (key_groups or {}).get(cand_key or "") if cand_key else None
    third_party_rucs = list(document.get("third_party_rucs") or [])
    secondary_fields = list(document.get("secondary_review_fields") or [])
    observed_issuer_name = original.get("issuer_name") or normalized.get("issuer_name") or None
    source_locator = document.get("display_location") or document.get("original_path")
    # El control aritmético de los importes, si se pudo hacer.
    arithmetic = next((
        check for check in document.get("derived_checks") or []
        if check.get("check") == "arithmetic_total_reconciliation"
        and check.get("method") == "observed_components"
    ), {})
    # Moneda declarada para la detracción, la retención y la percepción; sólo se usa si es una sola.
    settlement_currencies = {}
    for field in SETTLEMENT_FIELDS:
        codes = {
            item["decision_comment"].removeprefix("Moneda del importe: ")
            for item in evidence.get(field, [])
            if (item.get("decision_comment") or "").startswith("Moneda del importe: ")
        }
        if len(codes) == 1:
            settlement_currencies[f"moneda_{field}"] = next(iter(codes))

    return {
        # La fila: primero los importes y su cuadre; luego la denominación, la ubicación del original, los estados,
        # los datos del padrón y de la reconciliación; después los datos normalizados y, al final,
        # las columnas auxiliares de cada hoja.
        **settlement_currencies,
        **{
            f"importe_{field}": _numeric_or_none(normalized.get(field))
            for _, field in AMOUNT_COLUMNS
            # En bases antiguas subtotal podía significar solo la base gravada.
            if field != "subtotal" or arithmetic
        },
        "importe_calculado": _numeric_or_none(arithmetic.get("expected_total")),
        "importe_diferencia": _numeric_or_none(arithmetic.get("difference")),
        "estado_importes": arithmetic.get("status") or "requiere nueva lectura",
        "fuente_importes": (arithmetic.get("reading_source") or "Lectura de página") if arithmetic else None,
        # La denominación muestra el título observado; la línea de la que salió va aparte.
        **_document_title_fields(evidence, normalized.get("document_type")),
        "document_id": document.get("id"),
        "source_id": document.get("source_id"),
        "original_path": document.get("original_path"),
        "container_chain_str": " -> ".join(str(x) for x in chain),
        "member_path": document.get("member_path"),
        "nested_depth": document.get("nested_depth", 0),
        "display_location": source_locator,
        "source_format": document.get("source_format"),
        "reading_source": document.get("reading_source", "native_pdf_text"),
        "page_range": f"{document.get('page_start', 1)}-{document.get('page_end', 1)}",
        "block_range": f"{document.get('block_start', 1)}-{document.get('block_end', 1)}",
        "detection_status": det_status,
        "identity_status": id_status,
        "detection_reason": document.get("detection_reason") or document.get("decision_comment") or "",
        "detection_excerpt": document.get("detection_excerpt") or "",
        "identity_reason": document.get("identity_reason") or "",
        "identity_excerpt": document.get("identity_excerpt") or "",
        "candidate_fiscal_key": cand_key,
        "candidate_key_group": grp["label"] if grp else ("sin_clave_candidata" if is_cdp else "no_aplica"),
        "representaciones_misma_clave": grp["count"] if grp else (1 if is_cdp else 0),
        "ubicaciones_misma_clave": grp["locations_str"] if grp else (source_locator if is_cdp else ""),
        "key_reconciliation_status": document.get("key_reconciliation_status", "no_aplica"),
        "key_reconciliation_reason": document.get("key_reconciliation_reason", ""),
        "is_primary_representation": bool(document.get("is_primary_representation", False)),
        "eligible_for_valid_sheet": bool(document.get("eligible_for_valid_sheet", False)),
        "official_validation_status": document.get("official_validation_status", "no_aplica" if not is_cdp else "pendiente"),
        "official_validation_date": document.get("official_validation_date") or None,
        "official_validation_source": document.get("official_validation_source") or None,
        "official_validation_reference": document.get("official_validation_reference") or None,
        "official_validation_reason": document.get("official_validation_reason", ""),
        "official_validation_channel": document.get("official_validation_channel") or "sin_canal",
        "classification_rule": document.get("classification_rule", ""),
        "non_receipt_type": document.get("non_receipt_type"),
        "comprobante_detectado": is_cdp,
        "identidad_completa": id_completa,
        "renombrable": document.get("renombrable"),
        "issuer_role_status": document.get("issuer_role_status", "not_found"),
        "emitter_ruc_original": original.get("emitter_ruc"),
        "emitter_ruc_normalized": normalized.get("emitter_ruc"),
        "issuer_name_normalized": observed_issuer_name,
        "emitter_padron_status": document.get("emitter_padron_status", "no_evaluado"),
        "emitter_padron_razon_social": document.get("emitter_padron_razon_social"),
        "emitter_padron_estado": document.get("emitter_padron_estado"),
        "emitter_padron_cond_domicilio": document.get("emitter_padron_cond_domicilio"),
        "emitter_padron_cutoff_date": document.get("emitter_padron_cutoff_date"),
        "emitter_padron_source": document.get("emitter_padron_source"),
        "emitter_name_discrepancy": bool(document.get("emitter_name_discrepancy", False)),
        "recipient_role_status": document.get("recipient_role_status", "not_found"),
        "recipient_ruc_original": original.get("recipient_ruc"),
        "recipient_ruc_normalized": normalized.get("recipient_ruc"),
        "recipient_name_normalized": normalized.get("recipient_name") or original.get("recipient_name"),
        "recipient_ruc_match": document.get("recipient_ruc_match", False),
        "recipient_padron_status": document.get("recipient_padron_status", "no_evaluado"),
        "recipient_padron_razon_social": document.get("recipient_padron_razon_social"),
        "recipient_padron_estado": document.get("recipient_padron_estado"),
        "recipient_padron_cond_domicilio": document.get("recipient_padron_cond_domicilio"),
        "third_party_rucs_str": ", ".join(third_party_rucs),
        "document_type_original": original.get("document_type"),
        "document_type_normalized": normalized.get("document_type"),
        "series_original": original.get("series"),
        "series_normalized": normalized.get("series"),
        "correlativo_original": original.get("correlativo") or normalized.get("correlativo_original"),
        "correlativo_normalized": normalized.get("correlativo_normalized"),
        "issue_date_original": original.get("issue_date"),
        "issue_date_normalized": normalized.get("issue_date"),
        "due_date_normalized": normalized.get("due_date"),
        "currency_normalized": normalized.get("currency"),
        "payment_condition_normalized": normalized.get("payment_condition"),
        "referenced_document_normalized": normalized.get("referenced_document"),
        "subtotal_normalized": normalized.get("subtotal"),
        "inafecto_normalized": normalized.get("inafecto"),
        "exonerado_normalized": normalized.get("exonerado"),
        "otros_cargos_normalized": normalized.get("otros_cargos"),
        "igv_normalized": normalized.get("igv"),
        "total_original": original.get("total"),
        "total_normalized": normalized.get("total"),
        "general_description": document.get("general_description") or normalized.get("general_description"),
        "line_items_count": len(line_items),
        "line_items_summary": "; ".join(summary_parts),
        "secondary_review_fields_str": ", ".join(secondary_fields),
        "requires_human_review": document.get("requires_human_review", True),
        "decision_comment": document.get("decision_comment", ""),
        "line_items_json": _json(line_items),
        "derived_checks_json": _json(document.get("derived_checks") or []),
        "field_statuses_json": _json(document.get("field_statuses") or {}),
        "fields_original_json": _json(original),
        "fields_normalized_json": _json(normalized),
        "field_evidence_json": _json(evidence),
        "confidence_json": _json(document.get("confidence") or {}),
        "candidates_json": _json(document.get("candidates") or {}),
        "observations": "; ".join(document.get("observations") or []),
        "read_error": document.get("read_error", False),
        "read_error_details": "; ".join(document.get("read_error_details") or []),
        "text_sidecar_path": document.get("text_sidecar_path"),
        "normalized_filename": document.get("normalized_filename"),
        "organized_path": document.get("organized_path"),
        # Columnas de la hoja "CDP válidos"
        "valid_emitter_ruc": str(normalized.get("emitter_ruc")) if normalized.get("emitter_ruc") else None,
        "valid_issuer_name_observed": observed_issuer_name,
        "valid_document_type": str(normalized.get("document_type")) if normalized.get("document_type") else None,
        "valid_series": str(normalized.get("series")) if normalized.get("series") else None,
        "valid_correlativo": str(normalized.get("correlativo_normalized")) if normalized.get("correlativo_normalized") else None,
        "valid_issue_date": normalized.get("issue_date") or None,
        "valid_recipient_ruc": str(normalized.get("recipient_ruc")) if normalized.get("recipient_ruc") else None,
        "valid_currency": normalized.get("currency") or None,
        "valid_igv_numeric": _numeric_or_none(normalized.get("igv")),
        "valid_total_numeric": _numeric_or_none(normalized.get("total")),
        "valid_source_locator": source_locator,
        "valid_page_locator": _page_locator_str(document),
        "valid_description_observed": document.get("general_description") or normalized.get("general_description"),
        "detraccion_visible": visible_labels(evidence, "detraccion_visible"),
        "retencion_visible": visible_labels(evidence, "retencion_visible"),
        "buen_contribuyente_visible": visible_labels(evidence, "buen_contribuyente_visible"),
        "visible_tax_evidence": " | ".join(
            f"{page}: {fragment}" for field in (
                "detraccion_visible", "retencion_visible", "buen_contribuyente_visible"
            )
            for page, fragment in [_first_evidence_info(evidence, field, source_locator)]
            if fragment
        ),
    }


# Filas de la hoja notas: las notas de crédito y de débito, con el comprobante que modifican
# y si ese comprobante está en el mismo lote.
def _ajustes_vinculados_rows(
    documents: Sequence[dict[str, Any]],
    key_groups: Mapping[str, Mapping[str, Any]],
) -> list[dict[str, Any]]:
    # Las series se repiten entre empresas, así que cada nota se busca dentro de su propio emisor.
    by_serial: dict[tuple[str, str, str], list[str]] = {}
    for doc in documents:
        if normalize_detection_status(doc.get("detection_status")) != DETECTED:
            continue
        norm = doc.get("fields_normalized") or {}
        if norm.get("document_type") not in {"01", "03"}:
            continue
        issuer = str(norm.get("emitter_ruc") or "")
        ser = str(norm.get("series") or "").upper()
        corr = str(norm.get("correlativo_normalized") or "")
        if issuer and ser and corr:
            loc = f"{doc.get('display_location') or doc.get('original_path')} (pág. {doc.get('page_start', 1)}-{doc.get('page_end', 1)})"
            locations = by_serial.setdefault((issuer, ser, corr), [])
            if loc not in locations:
                locations.append(loc)

    rows: list[dict[str, Any]] = []
    for doc in documents:
        if normalize_detection_status(doc.get("detection_status")) != DETECTED:
            continue
        r = _row(doc, key_groups)
        doc_type = r["document_type_normalized"]
        if doc_type not in {"07", "08"} or r["key_reconciliation_status"] == "ajuste_concordante_secundario":
            continue
        evidence = doc.get("field_evidence") or {}
        ref_page, ref_excerpt = _first_evidence_info(evidence, "referenced_document", f"pág. {r['page_range']}")
        # Se busca el comprobante afectado dentro del lote, y dentro del mismo emisor.
        ref_doc = r["referenced_document_normalized"]
        affected_present = "Sin referencia leída"
        affected_locs = ""
        if ref_doc and "-" in str(ref_doc):
            parts = str(ref_doc).split("-", 1)
            ref_ser = parts[0].strip().upper()
            ref_corr = normalize_correlativo(parts[1].strip()).normalized or parts[1].strip()
            matched_locs = by_serial.get((r["emitter_ruc_normalized"] or "", ref_ser, ref_corr), [])
            if matched_locs:
                affected_present = "Sí"
                affected_locs = " | ".join(matched_locs)
            else:
                affected_present = "No encontrado en lote"

        rows.append(
            {
                **r,
                "note_type_label": "Nota de crédito" if doc_type == "07" else "Nota de débito",
                "referenced_document_page": ref_page if ref_doc else "-",
                "referenced_document_excerpt": ref_excerpt,
                "affected_present_in_batch": affected_present,
                "affected_locations_in_batch": affected_locs,
            }
        )
    return rows


# Filas de la hoja revision: todo lo que una persona debería mirar en el original.
# Cada hallazgo apunta a su archivo, su página y su fragmento.
def _review_rows(
    documents: Sequence[dict[str, Any]],
    inventory: Sequence[dict[str, Any]],
    key_groups: Mapping[str, Mapping[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    # Primero, las incidencias de archivos y de ZIP (cifrados, corruptos, bloqueados o con límite alcanzado).
    for inv in inventory:
        status = inv.get("inventory_status", "")
        if status in {"encrypted", "corrupted", "path_traversal_blocked", "limit_skipped", "read_error"}:
            chain = inv.get("container_chain") or []
            rows.append(
                {
                    "document_id": None,
                    "finding_type": f"incidencia_contenedor_{status}",
                    "priority": "alta",
                    "detection_status": status,
                    "identity_status": IDENTITY_NOT_APPLICABLE,
                    "review_field": "contenedor_o_archivo",
                    "observed_value": inv.get("detected_format") or status,
                    "verification_page": "-",
                    "evidence_excerpt": inv.get("status_reason") or "",
                    "review_reason": inv.get("status_reason") or f"Incidencia en contenedor/archivo: {status}",
                    "display_location": inv.get("display_location") or inv.get("source_path"),
                    "original_path": inv.get("source_path"),
                    "container_chain_str": " -> ".join(str(x) for x in chain),
                    "member_path": inv.get("member_path"),
                    "page_range": "-",
                    "candidate_fiscal_key": None,
                    "emitter_ruc": None,
                    "issuer_role_status": "n/a",
                    "recipient_ruc": None,
                    "recipient_role_status": "n/a",
                    "recipient_ruc_match": False,
                    "document_identity": "-",
                    "total_normalized": None,
                    "candidates_or_evidence": inv.get("action_taken") or "",
                }
            )

    for doc in documents:
        r = _row(doc, key_groups)
        evidence = doc.get("field_evidence") or {}
        original = doc.get("fields_original") or {}
        normalized = doc.get("fields_normalized") or {}
        default_page = f"pág. {r['page_range']}"
        doc_ident = f"{r['document_type_normalized'] or '?'}-{r['series_normalized'] or '?'}-{r['correlativo_normalized'] or r['correlativo_original'] or '?'}"

        # Documentos que no se pudieron leer.
        if r["detection_status"] == READ_ERROR or doc.get("read_error"):
            rows.append(
                {
                    "document_id": r["document_id"],
                    "finding_type": "error_de_lectura",
                    "priority": "alta",
                    "detection_status": r["detection_status"],
                    "identity_status": r["identity_status"],
                    "review_field": "lectura_pagina",
                    "observed_value": r["read_error_details"] or "ilegible",
                    "verification_page": default_page,
                    "evidence_excerpt": (doc.get("extracted_text") or "")[:160],
                    "review_reason": f"Error de lectura: {r['read_error_details'] or r['observations']}",
                    "display_location": r["display_location"],
                    "original_path": r["original_path"],
                    "container_chain_str": r["container_chain_str"],
                    "member_path": r["member_path"],
                    "page_range": r["page_range"],
                    "candidate_fiscal_key": r["candidate_fiscal_key"],
                    "emitter_ruc": r["emitter_ruc_normalized"] or r["emitter_ruc_original"],
                    "issuer_role_status": r["issuer_role_status"],
                    "recipient_ruc": r["recipient_ruc_normalized"] or r["recipient_ruc_original"],
                    "recipient_role_status": r["recipient_role_status"],
                    "recipient_ruc_match": r["recipient_ruc_match"],
                    "document_identity": doc_ident,
                    "total_normalized": r["total_normalized"],
                    "candidates_or_evidence": r["candidates_json"],
                }
            )
            continue

        # Documentos indeterminados: hay señales parciales, pero no bastan para decidir que son comprobantes.
        if r["detection_status"] == AMBIGUOUS:
            first_line = (doc.get("extracted_text") or "").splitlines()
            excerpt_str = first_line[0][:160] if first_line else ""
            rows.append(
                {
                    "document_id": r["document_id"],
                    "finding_type": "indeterminado",
                    "priority": "alta",
                    "detection_status": r["detection_status"],
                    "identity_status": r["identity_status"],
                    "review_field": "deteccion_comprobante",
                    "observed_value": doc_ident if doc_ident != "?-?-?" else "señales parciales",
                    "verification_page": default_page,
                    "evidence_excerpt": excerpt_str,
                    "review_reason": r["detection_reason"] or r["observations"] or r["decision_comment"],
                    "display_location": r["display_location"],
                    "original_path": r["original_path"],
                    "container_chain_str": r["container_chain_str"],
                    "member_path": r["member_path"],
                    "page_range": r["page_range"],
                    "candidate_fiscal_key": r["candidate_fiscal_key"],
                    "emitter_ruc": r["emitter_ruc_normalized"] or r["emitter_ruc_original"],
                    "issuer_role_status": r["issuer_role_status"],
                    "recipient_ruc": r["recipient_ruc_normalized"] or r["recipient_ruc_original"],
                    "recipient_role_status": r["recipient_role_status"],
                    "recipient_ruc_match": r["recipient_ruc_match"],
                    "document_identity": doc_ident,
                    "total_normalized": r["total_normalized"],
                    "candidates_or_evidence": r["candidates_json"],
                }
            )
            continue

        # De aquí en adelante, sólo los comprobantes detectados.
        if r["detection_status"] != DETECTED:
            continue

        # 1. Revisión de representaciones en conflicto por misma clave fiscal
        if r["key_reconciliation_status"] == "en_conflicto":
            rows.append(
                {
                    "document_id": r["document_id"],
                    "finding_type": "representaciones_en_conflicto",
                    "priority": "alta",
                    "detection_status": r["detection_status"],
                    "identity_status": r["identity_status"],
                    "review_field": "candidate_fiscal_key",
                    "observed_value": doc.get("candidate_fiscal_key") or doc_ident,
                    "verification_page": default_page,
                    "evidence_excerpt": r["identity_excerpt"] or r["detection_excerpt"],
                    "review_reason": r["key_reconciliation_reason"] or "Conflicto entre representaciones de la misma clave.",
                    "display_location": r["display_location"],
                    "original_path": r["original_path"],
                    "container_chain_str": r["container_chain_str"],
                    "member_path": r["member_path"],
                    "page_range": r["page_range"],
                    "candidate_fiscal_key": doc.get("candidate_fiscal_key"),
                    "emitter_ruc": r["emitter_ruc_normalized"] or r["emitter_ruc_original"],
                    "issuer_role_status": r["issuer_role_status"],
                    "recipient_ruc": r["recipient_ruc_normalized"] or r["recipient_ruc_original"],
                    "recipient_role_status": r["recipient_role_status"],
                    "recipient_ruc_match": r["recipient_ruc_match"],
                    "document_identity": doc_ident,
                    "total_normalized": r["total_normalized"],
                    "candidates_or_evidence": r["ubicaciones_misma_clave"],
                }
            )
        # 2. Revisión de campos de identidad faltantes, dudosos (O/0, I/1) o en conflicto interno
        elif r["identity_status"] != IDENTITY_COMPLETE:
            for id_field, label in (
                ("document_type", "document_type"),
                ("emitter_ruc", "emitter_ruc"),
                ("series", "series"),
                ("correlativo", "correlativo"),
            ):
                norm_key = "correlativo_normalized" if id_field == "correlativo" else id_field
                obs_val = original.get(id_field) or normalized.get(id_field)
                norm_val = normalized.get(norm_key)
                cands = (doc.get("candidates") or {}).get(id_field) or []
                ver_page, ver_excerpt = _first_evidence_info(evidence, id_field, default_page)
                if not ver_excerpt:
                    ver_excerpt = r["identity_excerpt"] or r["detection_excerpt"] or ""

                if len(cands) > 1:
                    rows.append(
                        {
                            "document_id": r["document_id"],
                            "finding_type": "identidad_en_conflicto",
                            "priority": "alta",
                            "detection_status": r["detection_status"],
                            "identity_status": r["identity_status"],
                            "review_field": label,
                            "observed_value": f"candidatos={cands} (literal={obs_val or 's/d'})",
                            "verification_page": ver_page,
                            "evidence_excerpt": ver_excerpt,
                            "review_reason": r["identity_reason"] or f"Candidatos múltiples en {label}: {cands}",
                            "display_location": r["display_location"],
                            "original_path": r["original_path"],
                            "container_chain_str": r["container_chain_str"],
                            "member_path": r["member_path"],
                            "page_range": r["page_range"],
                            "candidate_fiscal_key": r["candidate_fiscal_key"],
                            "emitter_ruc": r["emitter_ruc_normalized"] or r["emitter_ruc_original"],
                            "issuer_role_status": r["issuer_role_status"],
                            "recipient_ruc": r["recipient_ruc_normalized"] or r["recipient_ruc_original"],
                            "recipient_role_status": r["recipient_role_status"],
                            "recipient_ruc_match": r["recipient_ruc_match"],
                            "document_identity": doc_ident,
                            "total_normalized": r["total_normalized"],
                            "candidates_or_evidence": r["candidates_json"],
                        }
                    )
                elif not norm_val:
                    rows.append(
                        {
                            "document_id": r["document_id"],
                            "finding_type": "identidad_incompleta",
                            "priority": "alta",
                            "detection_status": r["detection_status"],
                            "identity_status": r["identity_status"],
                            "review_field": label,
                            "observed_value": obs_val or "no_observado",
                            "verification_page": ver_page,
                            "evidence_excerpt": ver_excerpt,
                            "review_reason": r["identity_reason"] or f"Campo de identidad {label} faltante o con carácter dudoso.",
                            "display_location": r["display_location"],
                            "original_path": r["original_path"],
                            "container_chain_str": r["container_chain_str"],
                            "member_path": r["member_path"],
                            "page_range": r["page_range"],
                            "candidate_fiscal_key": r["candidate_fiscal_key"],
                            "emitter_ruc": r["emitter_ruc_normalized"] or r["emitter_ruc_original"],
                            "issuer_role_status": r["issuer_role_status"],
                            "recipient_ruc": r["recipient_ruc_normalized"] or r["recipient_ruc_original"],
                            "recipient_role_status": r["recipient_role_status"],
                            "recipient_ruc_match": r["recipient_ruc_match"],
                            "document_identity": doc_ident,
                            "total_normalized": r["total_normalized"],
                            "candidates_or_evidence": r["candidates_json"],
                        }
                    )

        # 3. Revisión de receptor solo cuando el adquirente no fue identificado o su rol está en conflicto
        # (un adquirente distinto con RUC válido es resultado del filtro, no una alerta de revisión)
        if r["recipient_role_status"] in {"not_found", "candidate_unassigned"}:
            ver_page, ver_excerpt = _first_evidence_info(evidence, "recipient_ruc", default_page)
            rows.append(
                {
                    "document_id": r["document_id"],
                    "finding_type": "ruc_receptor_no_identificado",
                    "priority": "media",
                    "detection_status": r["detection_status"],
                    "identity_status": r["identity_status"],
                    "review_field": "recipient_ruc",
                    "observed_value": r["recipient_ruc_original"] or r["recipient_ruc_normalized"] or "no_observado",
                    "verification_page": ver_page,
                    "evidence_excerpt": ver_excerpt or r["detection_excerpt"],
                    "review_reason": "RUC adquirente sin determinar con rol explícito.",
                    "display_location": r["display_location"],
                    "original_path": r["original_path"],
                    "container_chain_str": r["container_chain_str"],
                    "member_path": r["member_path"],
                    "page_range": r["page_range"],
                    "candidate_fiscal_key": r["candidate_fiscal_key"],
                    "emitter_ruc": r["emitter_ruc_normalized"] or r["emitter_ruc_original"],
                    "issuer_role_status": r["issuer_role_status"],
                    "recipient_ruc": r["recipient_ruc_normalized"] or r["recipient_ruc_original"],
                    "recipient_role_status": r["recipient_role_status"],
                    "recipient_ruc_match": r["recipient_ruc_match"],
                    "document_identity": doc_ident,
                    "total_normalized": r["total_normalized"],
                    "candidates_or_evidence": r["candidates_json"],
                }
            )

        # La hoja de revisión señala tanto lecturas dudosas como diferencias de importes.
        for chk in doc.get("derived_checks") or []:
            if not chk.get("consistent", True):
                chk_name = chk.get("check")
                low_confidence = chk_name == "low_ocr_confidence"
                page_located = chk_name in {"missing_description", "missing_issue_date", "missing_issuer_name"}
                if low_confidence or page_located or (chk_name == "arithmetic_total_reconciliation" and chk.get("method") == "observed_components"):
                    if low_confidence or page_located:
                        ver_page, ver_excerpt = f"pág. {chk['page']}", chk["excerpt"]
                    else:
                        ver_page, ver_excerpt = _first_evidence_info(evidence, chk.get("compared_field", "total"), default_page)
                    rows.append(
                        {
                            "document_id": r["document_id"],
                            "finding_type": {
                                "low_ocr_confidence": "lectura_ocr_dudosa",
                                "missing_description": "descripcion_pendiente",
                                "missing_issue_date": "fecha_de_emision_ausente",
                                "missing_issuer_name": "razon_social_emisor_ausente",
                                "arithmetic_total_reconciliation": "revision_secundaria_importes",
                            }[chk_name],
                            "priority": "media",
                            "detection_status": r["detection_status"],
                            "identity_status": r["identity_status"],
                            "review_field": chk.get("field", "importes_totales"),
                            "observed_value": chk.get("observed_value", chk.get("status")),
                            "verification_page": ver_page,
                            "evidence_excerpt": ver_excerpt,
                            "review_reason": chk.get("reason") or (
                                f"Los componentes suman {chk.get('expected_total')} "
                                f"y el total leído es {chk.get('observed_total')}."
                            ),
                            "display_location": r["display_location"],
                            "original_path": r["original_path"],
                            "container_chain_str": r["container_chain_str"],
                            "member_path": r["member_path"],
                            "page_range": r["page_range"],
                            "candidate_fiscal_key": r["candidate_fiscal_key"],
                            "emitter_ruc": r["emitter_ruc_normalized"] or r["emitter_ruc_original"],
                            "issuer_role_status": r["issuer_role_status"],
                            "recipient_ruc": r["recipient_ruc_normalized"] or r["recipient_ruc_original"],
                            "recipient_role_status": r["recipient_role_status"],
                            "recipient_ruc_match": r["recipient_ruc_match"],
                            "document_identity": doc_ident,
                            "total_normalized": r["total_normalized"],
                            "candidates_or_evidence": _json(chk),
                        }
                    )
                elif chk_name == "line_item_arithmetic_reconciliation":
                    discs = chk.get("discrepancies") or []
                    rows.append(
                        {
                            "document_id": r["document_id"],
                            "finding_type": "revision_secundaria_detalle",
                            "priority": "media",
                            "detection_status": r["detection_status"],
                            "identity_status": r["identity_status"],
                            "review_field": "line_items",
                            "observed_value": "; ".join(str(x) for x in discs),
                            "verification_page": default_page,
                            "evidence_excerpt": "; ".join(str(x) for x in discs[:2]),
                            "review_reason": (
                                f"Discrepancia secundaria en cantidad × valor unitario de ítem(s); "
                                f"la detección ({r['detection_status']}) e identidad ({r['identity_status']}) no se alteran."
                            ),
                            "display_location": r["display_location"],
                            "original_path": r["original_path"],
                            "container_chain_str": r["container_chain_str"],
                            "member_path": r["member_path"],
                            "page_range": r["page_range"],
                            "candidate_fiscal_key": r["candidate_fiscal_key"],
                            "emitter_ruc": r["emitter_ruc_normalized"] or r["emitter_ruc_original"],
                            "issuer_role_status": r["issuer_role_status"],
                            "recipient_ruc": r["recipient_ruc_normalized"] or r["recipient_ruc_original"],
                            "recipient_role_status": r["recipient_role_status"],
                            "recipient_ruc_match": r["recipient_ruc_match"],
                            "document_identity": doc_ident,
                            "total_normalized": r["total_normalized"],
                            "candidates_or_evidence": _json(chk),
                        }
                    )
    return rows


# Filas de la hoja inventario: todos los archivos y miembros de ZIP encontrados con su estado,
# más una fila por cada bloque descartado por ser un soporte (HES, orden de compra, estado de cuenta...).
def _inventory_and_exclusions_rows(
    inventory: Sequence[dict[str, Any]],
    documents: Sequence[dict[str, Any]],
    pending_units: Sequence[str] = (),
) -> list[dict[str, Any]]:
    docs_by_source: dict[int, list[dict[str, Any]]] = {}
    for doc in documents:
        sid = doc.get("source_id")
        if sid is not None:
            docs_by_source.setdefault(int(sid), []).append(doc)

    # Unidades que quedaron sin procesar por una cancelación.
    pending_set = set(pending_units)
    rows: list[dict[str, Any]] = []
    for inv in inventory:
        sid = int(inv.get("source_id") or 0)
        chain = inv.get("container_chain") or []
        chain_str = " -> ".join(str(x) for x in chain)
        display_loc = inv.get("display_location") or inv.get("source_path") or ""
        inv_status = str(inv.get("inventory_status") or "")
        matched_docs = docs_by_source.get(sid, [])

        if matched_docs:
            statuses = sorted({normalize_detection_status(d.get("detection_status")) for d in matched_docs})
            det_status_str = ", ".join(statuses)
            non_receipt_types = sorted({str(d.get("non_receipt_type")) for d in matched_docs if d.get("non_receipt_type")})
            excl_type = ", ".join(non_receipt_types) if non_receipt_types else ("comprobante_de_pago" if DETECTED in statuses else inv_status)
            page_ranges = ", ".join(f"p{d.get('page_start', 1)}-{d.get('page_end', 1)}/b{d.get('block_start', 1)}-{d.get('block_end', 1)}" for d in matched_docs)
            action = "processed_and_classified"
        elif display_loc in pending_set:
            det_status_str = "pendiente_por_cancelacion"
            excl_type = "unidad_pendiente_por_cancelacion"
            page_ranges = "-"
            action = "pending_after_cancellation"
        else:
            det_status_str = "no_procesado" if inv_status != "supported" else "sin_resultado"
            excl_type = "xml_excluido_por_formato" if inv_status == "excluded_xml" else inv_status
            page_ranges = "-"
            action = str(inv.get("action_taken") or "")

        rows.append(
            {
                "source_id": sid,
                "record_kind": "inventario_fuente",
                "source_path": inv.get("source_path"),
                "container_chain_str": chain_str,
                "member_path": inv.get("member_path"),
                "nested_depth": inv.get("nested_depth", 0),
                "display_location": display_loc,
                "page_block_range": page_ranges,
                "detected_format": inv.get("detected_format"),
                "inventory_status": inv_status,
                "detection_status": det_status_str,
                "exclusion_or_support_type": excl_type,
                "status_reason": inv.get("status_reason") or "",
                "action_taken": action,
                "byte_size": inv.get("byte_size", 0),
                "compressed_size": inv.get("compressed_size"),
            }
        )

        # Agregar detalle explícito por cada bloque clasificado como no_cdp (HES, SAP, OC, guía, banco, SUNAT)
        for d in matched_docs:
            if normalize_detection_status(d.get("detection_status")) == NOT_RECEIPT:
                rows.append(
                    {
                        "source_id": sid,
                        "record_kind": "exclusion_soporte_no_cdp",
                        "source_path": d.get("original_path"),
                        "container_chain_str": chain_str,
                        "member_path": d.get("member_path"),
                        "nested_depth": d.get("nested_depth", 0),
                        "display_location": d.get("display_location") or display_loc,
                        "page_block_range": f"p{d.get('page_start', 1)}-{d.get('page_end', 1)} / b{d.get('block_start', 1)}-{d.get('block_end', 1)}",
                        "detected_format": d.get("source_format"),
                        "inventory_status": "excluded_non_cdp_support",
                        "detection_status": NOT_RECEIPT,
                        "exclusion_or_support_type": d.get("non_receipt_type") or "soporte_no_fiscal",
                        "status_reason": d.get("detection_reason") or d.get("decision_comment") or "",
                        "action_taken": "excluded_from_cdp_results",
                        "byte_size": inv.get("byte_size", 0),
                        "compressed_size": inv.get("compressed_size"),
                    }
                )
    return rows


# Escribe una hoja de Excel: encabezado, filas, formatos (texto, número, fecha), filtros y anchos de columna.
def _style_and_populate_sheet(
    sheet: Any,
    columns: Sequence[tuple[str, str]],
    rows: Sequence[Mapping[str, Any]],
    text_columns: set[str],
    numeric_columns: set[str] | None = None,
    table_name: str | None = None,
    normalizar_encabezados: bool = False,
) -> None:
    numeric_cols = numeric_columns or set()
    # Convierte «Razón social observada» en razon_social_observada: sin tildes, en minúsculas y con guiones bajos.
    def nombre_columna(label: str) -> str:
        texto = unicodedata.normalize("NFKD", label).encode("ascii", "ignore").decode("ascii").lower()
        return re.sub(r"[^a-z0-9]+", "_", texto).strip("_")

    sheet.append([nombre_columna(label) if normalizar_encabezados else label for label, _ in columns])
    # Encabezado con fondo azul y letras blancas.
    header_fill = PatternFill("solid", fgColor="174A70")
    for cell in sheet[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill

    # Una fila por registro. Las fechas en ISO se convierten en fechas reales de Excel.
    for row in rows:
        values = []
        for _, key in columns:
            value = row.get(key)
            if key in DATE_KEYS and isinstance(value, str):
                try:
                    value = date.fromisoformat(value)
                except ValueError:
                    # Una fecha antigua sin normalizar queda como se leyó.
                    pass
            values.append(value)
        sheet.append(values)
        row_number = sheet.max_row
        for column_idx, (label, key) in enumerate(columns, start=1):
            cell = sheet.cell(row_number, column_idx)
            value = values[column_idx - 1]
            if isinstance(value, str):
                # Los estados internos (en inglés) se muestran en español y con guiones bajos.
                traducciones = {
                    "native_pdf_text": "texto_digital", "ocr": "ocr", "not_found": "no_encontrado",
                    "candidate_unassigned": "candidato_sin_asignar", "verified_recipient": "receptor_identificado",
                    "verified_issuer": "emisor_identificado", "mismatch": "ruc_diferente",
                    "supported": "admitido", "excluded_xml": "xml_excluido",
                    "processed": "procesado", "processed_and_classified": "procesado_y_clasificado",
                    "skipped": "omitido", "read_error": "error_de_lectura",
                    "archive_limit_reached": "limite_del_archivo_alcanzado",
                    "archive_error": "error_de_archivo_comprimido",
                    "archive_traversal_blocked": "ruta_bloqueada",
                    "archive_encrypted": "archivo_comprimido_cifrado",
                    "container": "contenedor", "directory": "carpeta",
                    "error_recorded": "error_registrado",
                    "inspected_container": "contenedor_inspeccionado",
                    "inspected_directory": "carpeta_inspeccionada",
                    "skipped_limit": "omitido_por_limite",
                    "blocked_security": "bloqueado_por_seguridad",
                    "unsupported": "formato_no_admitido",
                    "excluded_from_cdp_results": "excluido_de_comprobantes",
                    "ignored_unsupported": "formato_no_admitido_omitido",
                    "excluded_by_format": "excluido_por_formato",
                    "requiere nueva lectura": "requiere_nueva_lectura",
                    "excluded_non_cdp_support": "soporte_excluido",
                    "pending_after_cancellation": "pendiente_por_cancelacion",
                    "found": "encontrado", "unknown": "desconocido",
                }
                # Sólo se traducen las columnas de estado, no los textos libres.
                if label in {"Lectura", "Fuente de lectura", "Estado inventario", "Acción tomada", "Rol emisor", "Rol receptor", "Estado identidad", "Clasificación soporte / Exclusión", "Cuadre de importes"}:
                    value = traducciones.get(value, value)
                if label == "Fuente del cuadre":
                    value = {"Lectura de página": "lectura_de_pagina", "Recuadro de importes": "recuadro_de_importes"}.get(value, value)
                cell.value = value
                # El texto observado puede empezar por '=' y debe quedar literal.
                cell.data_type = "s"
            # Formato de la celda: texto (para no perder ceros), cifra con dos decimales o fecha.
            if label in text_columns and value is not None:
                cell.number_format = "@"
            elif label in numeric_cols and isinstance(value, (int, float)):
                cell.number_format = "#,##0.00"
            elif key in DATE_KEYS and isinstance(value, (date, datetime)):
                cell.number_format = DATE_FORMAT

    # Se deja fija la fila de encabezados, se pone un filtro (o una tabla) y más abajo se ajusta el ancho de las columnas.
    sheet.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(columns))
    if table_name:
        if sheet.max_row < 2:
            sheet.append([None] * len(columns))
        table_ref = f"A1:{last_col_letter}{sheet.max_row}"
        table = Table(displayName=table_name, ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)
    else:
        sheet.auto_filter.ref = f"A1:{last_col_letter}{max(sheet.max_row, 1)}"

    for index, (label, _) in enumerate(columns, start=1):
        width = min(max(len(label) + 4, 14), 44)
        if label in {
            "Archivo de origen",
            "Razón social observada",
            "Archivo origen",
            "Ubicación completa",
            "Ruta de salida",
            "Evidencia por campo JSON",
            "Glosa / Descripción general",
            "Resumen de ítems",
            "Motivo de revisión / Comentario",
            "Fragmento evidencia",
            "Fragmento detección",
            "Fragmento identidad",
            "Ubicaciones misma clave",
        }:
            width = 42
        sheet.column_dimensions[get_column_letter(index)].width = width


# Guarda el libro primero en un archivo temporal y después lo reemplaza, para no dañar el anterior
# si Excel lo tiene abierto.
def save_workbook(workbook: Workbook, path: Path) -> Path:
    # Excel puede tener abierto el libro anterior. Si el reemplazo falla, lo dejamos intacto.
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as handle:
        temporary = Path(handle.name)
    try:
        workbook.save(temporary)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
    return path.resolve()


# Punto de entrada: arma CDP.xlsx a partir de la base de la corrida.
# Hojas, en este orden: cdp_detectados, notas, revision, inventario y resumen (siempre al final).
def export_results(
    store: StateStore,
    output_root: str | Path,
    *,
    summary: ScanSummary | None = None,
) -> Path:
    # Paso 1: leer los documentos de la base. Al reexportar una corrida antigua se aplican las exigencias de hoy:
    # que el título del tipo documental esté a la vista y que la verificación oficial siempre quede «pendiente».
    documents = store.list_documents()
    for doc in documents:
        if normalize_detection_status(doc.get("detection_status")) == DETECTED:
            document_type = (doc.get("fields_normalized") or {}).get("document_type")
            titles = (doc.get("field_evidence") or {}).get("document_type") or []
            if not any(has_document_title(item.get("excerpt") or "", document_type) for item in titles):
                # Al reexportar una corrida antigua aplicamos también la exigencia de título.
                doc.update(
                    detection_status=AMBIGUOUS,
                    comprobante_detectado=False,
                    identidad_completa=False,
                    identity_status=IDENTITY_NOT_APPLICABLE,
                    detection_reason="La lectura guardada carece de evidencia del título del tipo documental.",
                )
        # Una corrida guardada tampoco puede acreditar una emisión desde esta aplicación.
        is_detected = normalize_detection_status(doc.get("detection_status")) == DETECTED
        doc.update(
            official_validation_status="pendiente" if is_detected else "no_aplica",
            official_validation_date=None,
            official_validation_reason="Esta aplicación no consulta la validez de emisión." if is_detected else "",
            eligible_for_valid_sheet=False,
        )

    # Paso 2: asegurar la reconciliación de las copias y armar los grupos por clave fiscal.
    _ensure_reconciled_documents(documents)
    inventory = store.list_inventory()
    key_groups = _build_candidate_key_groups(documents)
    # Paso 3: separar los comprobantes detectados. Las notas (07 y 08) van a su propia hoja.
    detected = [
        doc for doc in documents
        if normalize_detection_status(doc.get("detection_status"), read_error=bool(doc.get("read_error"))) == DETECTED
    ]
    cdp = [doc for doc in detected if (doc.get("fields_normalized") or {}).get("document_type") not in {"07", "08"}]
    # Mostramos una sola copia concordante. Si la clave tiene datos enfrentados,
    # conservamos las ubicaciones en Revisión para que se puedan cotejar con los originales.
    detected_rows = [
        _row(doc, key_groups) for doc in cdp
        if doc.get("key_reconciliation_status") != "concordante_secundario"
    ]
    detected_rows.sort(key=lambda row: (
        row["valid_emitter_ruc"] or "", row["valid_series"] or "",
        row["valid_correlativo"] or "", row["valid_source_locator"] or "",
    ))
    # Paso 4: preparar las filas de las otras hojas: notas, revisión e inventario.
    adjustments = _ajustes_vinculados_rows(documents, key_groups)
    review = _review_rows(documents, inventory, key_groups)
    run = store.get_latest_run_metadata() or {}
    run_details = run.get("run_config") or {}
    pending_units = summary.pending_units if summary else run_details.get("pending_units") or []
    inventory_rows = _inventory_and_exclusions_rows(inventory, documents, pending_units)
    status = summary.run_status if summary else run_details.get("run_status", "sin estado guardado")

    # Paso 5: columnas de la hoja principal. Se omite la fecha de verificación y se añaden identidad,
    # copias, lectura, padrón y evidencias.
    detected_columns = [
        *(column for column in CDP_VALIDOS_COLUMNS if column[1] != "official_validation_date"),
        ("Identidad", "identity_status"),
        ("Cotejo de duplicados", "key_reconciliation_status"),
        ("Representaciones", "representaciones_misma_clave"),
        ("Coincide RUC objetivo", "recipient_ruc_match"),
        ("Lectura", "reading_source"),
        ("Razón social del padrón", "emitter_padron_razon_social"),
        ("Fecha de corte del padrón", "emitter_padron_cutoff_date"),
        ("Buen contribuyente visible", "buen_contribuyente_visible"),
        ("Evidencia de leyendas", "visible_tax_evidence"),
        ("Evidencia de la denominación", "document_title_evidence"),
    ]
    # Paso 6: indicadores de la hoja resumen. «CDP confirmados oficialmente» es siempre 0 porque no hay consulta a SUNAT.
    metrics = [
        ("Estado de corrida", status),
        ("RUC receptor objetivo", run.get("target_recipient_ruc")),
        ("Representaciones CDP detectadas", len(cdp)),
        ("Filas en CDP detectados", len(detected_rows)),
        ("CDP confirmados oficialmente", 0),
        ("Notas de crédito y débito", len(adjustments)),
        ("Hallazgos para revisar", len(review)),
        ("Canal oficial de validación", "Sin conexión a consulta oficial"),
    ]
    # Paso 7: crear el libro y escribir las hojas (sin las columnas de soporte que no hacen falta) y, al final, el resumen.
    workbook = Workbook()
    workbook.remove(workbook.active)
    text_columns = {"RUC emisor", "RUC adquirente", "Tipo", "Serie", "Correlativo", "Página"}
    # Las columnas de soporte se reducen a las necesarias para volver al original.
    for title, columns, rows in (
        ("cdp_detectados", detected_columns, detected_rows),
        ("notas", [column for column in AJUSTES_COLUMNS if column[1] not in {
            "document_id", "original_path", "container_chain_str", "member_path",
            "detection_status", "candidate_fiscal_key", "note_type_label",
        }], adjustments),
        ("revision", [column for column in REVIEW_COLUMNS if column[1] not in {
            "document_id", "original_path", "container_chain_str", "member_path",
            "recipient_ruc_match", "candidates_or_evidence",
        }], review),
        ("inventario", [column for column in INVENTORY_COLUMNS if column[1] not in {
            "source_id", "source_path", "container_chain_str", "member_path", "nested_depth",
            "byte_size", "compressed_size",
        }], inventory_rows),
    ):
        _style_and_populate_sheet(
            workbook.create_sheet(title), columns, rows, text_columns,
            numeric_columns={label for label, key in (*AMOUNTS_EXPORT, *SETTLEMENT_EXPORT) if key.startswith("importe_")},
            normalizar_encabezados=True,
        )
    resumen = workbook.create_sheet("resumen")
    _style_and_populate_sheet(
        resumen, [("Indicador", "metric"), ("Valor", "value")],
        [{"metric": label, "value": value} for label, value in metrics], set(),
        normalizar_encabezados=True,
    )
    # Paso 8: guardar CDP.xlsx en la carpeta de salida.
    return save_workbook(workbook, Path(output_root) / OUTPUT_FILENAME)


