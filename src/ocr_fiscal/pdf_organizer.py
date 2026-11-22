"""Organizador de comprobantes en carpetas PDF por proveedor listas para archivo contable o fiscalización."""

from __future__ import annotations

import io
from pathlib import Path
import re
import shutil
import unicodedata
from typing import Any, Mapping, Sequence

import pymupdf
from PIL import Image, ImageOps

try:
    import pillow_heif  # type: ignore[import-not-found]

    pillow_heif.register_heif_opener()
except Exception:
    pass

from .archives import WorkUnit
from .models import DETECTED, IDENTITY_COMPLETE, NOT_RECEIPT


# Caracteres que Windows y otros sistemas no admiten en nombres de carpeta o archivo.
_INVALID_PATH_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')
# Formatos gráficos que se convierten a PDF para armar el expediente uniforme.
_IMAGE_FORMATS = {"jpg", "jpeg", "png", "tif", "tiff", "bmp", "webp", "heic"}
# Tope práctico de peso por archivo en Mesa de Partes Virtual (MPV) de SUNAT.
_SUNAT_FILE_LIMIT_MB = 10.0


def _sanitize_segment(value: str) -> str:
    """Limpia texto y reemplaza espacios por guiones bajos para rutas sin espacios."""
    ascii_text = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    cleaned = _INVALID_PATH_CHARS.sub(" ", ascii_text)
    cleaned = re.sub(r"\s+", "_", cleaned).strip(" ._")
    cleaned = re.sub(r"_+", "_", cleaned)
    return cleaned


def shorten_provider_name(name: str | None, max_len: int = 45) -> str:
    """Acorta una razón social manteniendo palabras enteras y uniéndolas sin espacios."""
    ascii_text = unicodedata.normalize("NFKD", name or "").encode("ascii", "ignore").decode("ascii")
    words_clean = _INVALID_PATH_CHARS.sub(" ", ascii_text)
    words_clean = re.sub(r"[.,;()\[\]]+", " ", words_clean)
    words_clean = re.sub(r"\s+", " ", words_clean).strip(" -_")
    if not words_clean:
        return ""
    if len(words_clean) > max_len:
        cut = words_clean[:max_len].rstrip(" -_")
        if " " in cut:
            word_cut = cut.rsplit(" ", 1)[0].strip(" -_")
            if len(word_cut) >= max(10, int(max_len * 0.55)):
                cut = word_cut
        words_clean = cut
    return re.sub(r"\s+", "_", words_clean).strip("_")


def format_provider_folder_name(ruc: str, provider_name: str | None, max_name_len: int = 45) -> str:
    """Arma la carpeta del proveedor sin espacios: `<RUC>-<Nombre_acortado>` o solo `<RUC>`."""
    safe_ruc = _sanitize_segment(ruc) or "SIN_RUC"
    short_name = shorten_provider_name(provider_name, max_len=max_name_len)
    if short_name:
        return f"{safe_ruc}-{short_name}"
    return safe_ruc


def _resolve_provider_labels(outcomes: Sequence[Any], max_name_len: int = 45) -> dict[str, str]:
    """Asigna una carpeta única y estable por cada RUC emisor de la corrida."""
    candidates_by_ruc: dict[str, list[tuple[int, str]]] = {}
    for item in outcomes:
        ruc = str(item.fields.get("emitter_ruc") or item.fields_original.get("emitter_ruc") or "").strip()
        if not ruc:
            continue
        padron_name = getattr(item, "emitter_padron_razon_social", None)
        obs_name = item.fields.get("issuer_name") or item.fields_original.get("issuer_name")
        if padron_name and str(padron_name).strip():
            candidates_by_ruc.setdefault(ruc, []).append((2, str(padron_name).strip()))
        elif obs_name and str(obs_name).strip():
            candidates_by_ruc.setdefault(ruc, []).append((1, str(obs_name).strip()))

    labels: dict[str, str] = {}
    for ruc, opts in candidates_by_ruc.items():
        # Prioriza la razón social oficial del padrón si existe; si no, la observada más frecuente/larga.
        opts.sort(key=lambda pair: (pair[0], len(pair[1]), pair[1]), reverse=True)
        labels[ruc] = format_provider_folder_name(ruc, opts[0][1], max_name_len=max_name_len)
    return labels


def _write_pdf_subset(
    unit: WorkUnit,
    page_start: int,
    page_end: int,
    total_blocks_in_unit: int,
    dest: Path,
    *,
    shares_physical_page: bool = False,
) -> tuple[bool, str | None]:
    """Copia un PDF entero intacto o extrae su rango de páginas exacto con PyMuPDF."""
    raw_bytes = unit.read_bytes()
    src_pdf = pymupdf.open(stream=raw_bytes, filetype="pdf") if raw_bytes is not None else pymupdf.open(unit.disk_path)
    with src_pdf:
        p1 = max(1, min(page_start, src_pdf.page_count))
        p2 = max(p1, min(page_end, src_pdf.page_count))
        is_full_untouched = (p1 == 1 and p2 == src_pdf.page_count and total_blocks_in_unit == 1)
        if is_full_untouched:
            if raw_bytes is not None:
                dest.write_bytes(raw_bytes)
            elif unit.disk_path is not None:
                shutil.copy2(unit.disk_path, dest)
            return True, None
        out_pdf = pymupdf.open()
        try:
            out_pdf.insert_pdf(src_pdf, from_page=p1 - 1, to_page=p2 - 1)
            out_pdf.save(dest, deflate=True, garbage=3)
        finally:
            out_pdf.close()
    return False, ("pagina_compartida_con_otro_bloque" if shares_physical_page else None)


def _build_images_pdf_bytes(
    unit: WorkUnit,
    page_start: int,
    page_end: int,
    rotation_by_page: Mapping[int, int],
    *,
    jpeg_quality: int = 92,
    max_side_px: int | None = None,
) -> tuple[bytes, int]:
    """Construye los bytes de un PDF a partir de una imagen (o TIFF multipágina)."""
    raw_bytes = unit.read_bytes()
    img = Image.open(io.BytesIO(raw_bytes)) if raw_bytes is not None else Image.open(unit.disk_path)
    out_pdf = pymupdf.open()
    exported_pages = 0
    try:
        with img:
            frame_count = max(1, int(getattr(img, "n_frames", 1)))
            p1 = max(1, min(page_start, frame_count))
            p2 = max(p1, min(page_end, frame_count))
            for page_no in range(p1, p2 + 1):
                img.seek(page_no - 1)
                frame = img.copy()
                try:
                    frame = ImageOps.exif_transpose(frame) or frame
                except Exception:
                    pass
                angle = int(rotation_by_page.get(page_no, 0))
                if angle in {90, 180, 270}:
                    frame = frame.rotate(-angle, expand=True)
                rgb = frame.convert("RGB")
                img_dpi = img.info.get("dpi")
                eff_dpi = (
                    float(img_dpi[0])
                    if isinstance(img_dpi, tuple) and len(img_dpi) == 2 and img_dpi[0] and float(img_dpi[0]) >= 72.0
                    else 150.0
                )
                page_w = max(72.0, rgb.width * 72.0 / eff_dpi)
                page_h = max(72.0, rgb.height * 72.0 / eff_dpi)
                if max_side_px is not None and max(rgb.width, rgb.height) > max_side_px:
                    ratio = max_side_px / float(max(rgb.width, rgb.height))
                    rgb = rgb.resize(
                        (max(1, int(rgb.width * ratio)), max(1, int(rgb.height * ratio))),
                        Image.Resampling.LANCZOS,
                    )
                buf = io.BytesIO()
                rgb.save(buf, format="JPEG", quality=jpeg_quality, optimize=True)
                pdf_page = out_pdf.new_page(width=page_w, height=page_h)
                pdf_page.insert_image(pdf_page.rect, stream=buf.getvalue())
                exported_pages += 1
        return out_pdf.tobytes(deflate=True, garbage=3), exported_pages
    finally:
        out_pdf.close()


def _write_images_as_pdf(
    unit: WorkUnit,
    page_start: int,
    page_end: int,
    rotation_by_page: Mapping[int, int],
    dest: Path,
) -> tuple[int, bool]:
    """Convierte una imagen (o cuadros de un TIFF multipágina) en PDF respetando en lo posible el tope MPV de 10 MB."""
    limit_bytes = int(_SUNAT_FILE_LIMIT_MB * 1024 * 1024)
    attempts = ((92, None), (80, None), (72, 2000), (68, 1600))
    chosen_bytes = b""
    pages = 0
    compressed = False
    for idx, (q, max_px) in enumerate(attempts):
        chosen_bytes, pages = _build_images_pdf_bytes(
            unit, page_start, page_end, rotation_by_page, jpeg_quality=q, max_side_px=max_px
        )
        if idx > 0:
            compressed = True
        if len(chosen_bytes) <= limit_bytes:
            break
    dest.write_bytes(chosen_bytes)
    return pages, compressed


def _write_pdf_pages_list(
    unit: WorkUnit,
    page_ranges: Sequence[tuple[int, int]],
    dest: Path,
) -> int:
    """Extrae uno o varios rangos de páginas de un PDF y los une en `dest`."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    raw_bytes = unit.read_bytes()
    src_pdf = pymupdf.open(stream=raw_bytes, filetype="pdf") if raw_bytes is not None else pymupdf.open(unit.disk_path)
    out_pdf = pymupdf.open()
    exported_pages = 0
    try:
        with src_pdf:
            for p_start, p_end in page_ranges:
                p1 = max(1, min(p_start, src_pdf.page_count))
                p2 = max(p1, min(p_end, src_pdf.page_count))
                out_pdf.insert_pdf(src_pdf, from_page=p1 - 1, to_page=p2 - 1)
                exported_pages += (p2 - p1 + 1)
        out_pdf.save(dest, deflate=True, garbage=3)
    finally:
        out_pdf.close()
    return exported_pages


def _export_outcome_to_pdf(
    unit: WorkUnit,
    source_format: str,
    page_start: int,
    page_end: int,
    total_blocks_in_unit: int,
    rotation_by_page: Mapping[int, int],
    dest: Path,
    *,
    shares_physical_page: bool = False,
) -> tuple[bool, str | None, int]:
    """Exporta un segmento a `dest` en formato PDF y devuelve `(exito, nota_opcional, paginas)`."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    fmt = source_format.lower()
    if fmt == "pdf":
        _, note = _write_pdf_subset(unit, page_start, page_end, total_blocks_in_unit, dest, shares_physical_page=shares_physical_page)
        pages = max(1, page_end - page_start + 1)
        return True, note, pages
    if fmt in _IMAGE_FORMATS:
        pages, compressed = _write_images_as_pdf(unit, page_start, page_end, rotation_by_page, dest)
        note = "imagen_recompresion_adaptativa_mpv" if compressed else None
        return True, note, pages
    return False, f"formato_no_convertible_a_pdf:{fmt}", 0


def _safe_dest_path(base_dest: Path, max_total_len: int = 235) -> Path:
    """Acorta el nombre de carpeta del proveedor si la ruta completa supera el límite seguro de Windows."""
    if len(str(base_dest)) <= max_total_len:
        return base_dest
    excess = len(str(base_dest)) - max_total_len
    parent = base_dest.parent
    new_folder_name = parent.name[: max(15, len(parent.name) - excess)].rstrip(" _-")
    return parent.parent / (new_folder_name or parent.name[:15]) / base_dest.name


def _set_pdf_derived_info(
    item: Any,
    *,
    size_mb: float,
    pages: int,
    over_limit: bool,
    linked_supports: Sequence[str] = (),
) -> None:
    item.derived_checks = [
        chk for chk in (getattr(item, "derived_checks", []) or []) if chk.get("check") != "pdf_expediente_info"
    ]
    item.derived_checks.append(
        {
            "check": "pdf_expediente_info",
            "size_mb": round(size_mb, 2),
            "pages": pages,
            "over_limit": over_limit,
            "linked_supports": list(linked_supports),
        }
    )


def _has_blocking_expediente_issue(item: Any, target_recipient_ruc: str | None) -> tuple[bool, str | None]:
    """Determina si un CDP completo NO debe ingresar a `01_cdp_por_proveedor/` por riesgo directo de reparo."""
    # La venta propia va primero: en una venta propia el receptor siempre es «otro».
    if target_recipient_ruc and str(item.fields.get("emitter_ruc") or "").strip() == target_recipient_ruc:
        return True, "VENTA_PROPIA_"
    if item.recipient_role_status == "mismatch":
        return True, "TERCERO_"
    if any(
        chk.get("check") == "sunat_qr_crosscheck" and chk.get("status") == "difiere"
        for chk in (getattr(item, "derived_checks", []) or [])
    ):
        return True, "QR_DIFIERE_"
    return False, None


def organize_pdf_expediente(
    output_root: Path,
    outcomes: Sequence[Any],
    units_by_source_id: Mapping[int, WorkUnit],
    source_formats_by_doc: Mapping[int, str],
    rotations_by_source_id: Mapping[int, Mapping[int, int]],
    unit_group_counts: Mapping[int, int],
    *,
    max_provider_name_len: int = 45,
    unify_notes_by_provider: bool = True,
    target_recipient_ruc: str | None = None,
) -> dict[str, int]:
    """Clasifica y convierte los comprobantes a PDF organizados por proveedor dentro de `expediente_pdf/`."""
    expediente_root = output_root / "expediente_pdf"
    expediente_root.mkdir(parents=True, exist_ok=True)
    provider_folders = _resolve_provider_labels(outcomes, max_name_len=max_provider_name_len)

    stats = {
        "total_exported": 0,
        "cdp_complete": 0,
        "notes_complete": 0,
        "pending_review": 0,
        "supports": 0,
        "linked_supports": 0,
        "secondary_duplicates": 0,
        "collisions": 0,
        "over_limit_mb": 0,
        "errors": 0,
    }

    # Precalculamos los vínculos físicos dentro del mismo archivo PDF:
    # si un PDF trae un comprobante apto (p. ej. página 1) junto con sus anexos (guía, OC o constancia SPOT),
    # exportamos esos anexos al lado del comprobante en la carpeta del proveedor en vez de dispersarlos.
    items_by_source: dict[int, list[Any]] = {}
    for item in outcomes:
        if item.source_id is not None:
            items_by_source.setdefault(item.source_id, []).append(item)

    support_owner_by_doc_id: dict[int, Any] = {}
    supports_by_owner_doc_id: dict[int, list[Any]] = {}
    for sid, group in items_by_source.items():
        if (source_formats_by_doc.get(group[0].document_id) or "").lower() != "pdf":
            continue
        ordered_group = sorted(group, key=lambda it: (it.page_start, it.block_start, it.document_id))
        eligible_primaries = [
            it for it in ordered_group
            if (
                it.detection_status == DETECTED
                and it.identity_status == IDENTITY_COMPLETE
                and bool(it.candidate_fiscal_key)
                and bool(it.is_primary_representation)
                and not it.read_error
                and not _has_blocking_expediente_issue(it, target_recipient_ruc)[0]
            )
        ]
        if not eligible_primaries:
            continue
        if len(eligible_primaries) == 1:
            owner = eligible_primaries[0]
            for it in ordered_group:
                if it.detection_status == NOT_RECEIPT and it.non_receipt_type != "xml_excluido_por_formato":
                    support_owner_by_doc_id[it.document_id] = owner
                    supports_by_owner_doc_id.setdefault(owner.document_id, []).append(it)
        else:
            last_owner: Any = None
            for it in ordered_group:
                if it in eligible_primaries:
                    last_owner = it
                elif (
                    last_owner is not None
                    and it.detection_status == NOT_RECEIPT
                    and it.non_receipt_type != "xml_excluido_por_formato"
                ):
                    support_owner_by_doc_id[it.document_id] = last_owner
                    supports_by_owner_doc_id.setdefault(last_owner.document_id, []).append(it)

    exported_dest_by_doc_id: dict[int, Path] = {}

    for item in outcomes:
        # Los soportes vinculados al mismo PDF de un CDP principal se exportan consolidados junto al CDP dueño.
        if item.document_id in support_owner_by_doc_id:
            continue

        unit = units_by_source_id.get(item.source_id) if item.source_id is not None else None
        if unit is None or item.read_error:
            if unit is None:
                continue
        src_fmt = (source_formats_by_doc.get(item.document_id) or unit.source_ref.detected_format or "").lower()
        orig_stem = _sanitize_segment(Path(unit.source_ref.member_path or item.original_path).stem) or f"doc_{item.document_id}"
        orig_stem = orig_stem[:40]
        doc_type = str(item.fields.get("document_type") or "")
        emitter_ruc = str(item.fields.get("emitter_ruc") or item.fields_original.get("emitter_ruc") or "").strip()
        provider_dir = provider_folders.get(emitter_ruc) or format_provider_folder_name(
            emitter_ruc or "SIN_RUC", getattr(item, "emitter_padron_razon_social", None) or item.fields.get("issuer_name"),
            max_name_len=max_provider_name_len,
        )

        collision_detected = False
        is_complete_primary = (
            item.detection_status == DETECTED
            and item.identity_status == IDENTITY_COMPLETE
            and bool(item.candidate_fiscal_key)
            and bool(item.is_primary_representation)
            and not item.read_error
        )
        blocked_issue, issue_prefix = _has_blocking_expediente_issue(item, target_recipient_ruc) if is_complete_primary else (False, None)

        if is_complete_primary and not blocked_issue:
            fiscal_stem = _sanitize_segment(str(item.candidate_fiscal_key))
            filename = f"{fiscal_stem}.pdf"
            if unify_notes_by_provider or doc_type not in {"07", "08"}:
                base_folder = "01_cdp_por_proveedor"
            else:
                base_folder = "02_notas_por_proveedor"
            candidate_dest = _safe_dest_path(expediente_root / base_folder / provider_dir / filename)
            if candidate_dest.exists():
                collision_detected = True
                dest = expediente_root / "03_revision_pendiente" / f"COLISION_{item.document_id:04d}_{filename}"
                category = "pending_review"
            else:
                dest = candidate_dest
                category = "notes_complete" if doc_type in {"07", "08"} else "cdp_complete"
        elif is_complete_primary and blocked_issue:
            fiscal_stem = _sanitize_segment(str(item.candidate_fiscal_key))
            dest = expediente_root / "03_revision_pendiente" / f"{issue_prefix}{item.document_id:04d}_{fiscal_stem}.pdf"
            category = "pending_review"
            item.requires_human_review = True
        elif item.detection_status == DETECTED and item.key_reconciliation_status in {
            "concordante_secundario",
            "ajuste_concordante_secundario",
        }:
            key_part = _sanitize_segment(str(item.candidate_fiscal_key or orig_stem))
            filename = f"DUP_{item.document_id:04d}_{key_part}.pdf"
            dest = _safe_dest_path(expediente_root / "05_duplicados_secundarios" / provider_dir / filename)
            category = "secondary_duplicates"
        elif item.detection_status == NOT_RECEIPT:
            sup_kind = _sanitize_segment(str(item.non_receipt_type or "soporte"))[:28]
            sup_ref = _sanitize_segment(str((item.fields or {}).get("support_reference") or ""))[:24]
            ref_part = f"_{sup_ref}" if sup_ref else ""
            filename = f"DOC_{item.document_id:04d}_{sup_kind}{ref_part}_{orig_stem}_p{item.page_start}-{item.page_end}.pdf"
            dest = expediente_root / "04_soportes_no_cdp" / filename
            category = "supports"
        else:
            filename = f"DOC_{item.document_id:04d}_{orig_stem}_p{item.page_start}-{item.page_end}.pdf"
            dest = expediente_root / "03_revision_pendiente" / filename
            category = "pending_review"

        # Evitamos cualquier sobrescritura accidental en carpetas de revisión, soportes o duplicados.
        if dest.exists():
            dest = dest.with_name(f"{dest.stem}_{item.document_id}.pdf")

        rot_map = rotations_by_source_id.get(item.source_id or -1, {})
        blocks_count = int(unit_group_counts.get(item.document_id, 1))
        shares_physical = any(
            other.document_id != item.document_id
            and not (other.page_end < item.page_start or other.page_start > item.page_end)
            for other in items_by_source.get(item.source_id or -1, ())
        )
        try:
            ok, note, exported_pages = _export_outcome_to_pdf(
                unit,
                src_fmt,
                item.page_start,
                item.page_end,
                blocks_count,
                rot_map,
                dest,
                shares_physical_page=shares_physical,
            )
        except Exception as exc:
            item.observations.append(f"error_exportacion_pdf: {exc}")
            stats["errors"] += 1
            continue

        if not ok:
            if note:
                item.observations.append(note)
            continue

        rel_path = dest.relative_to(output_root).as_posix()
        item.normalized_filename = dest.name
        item.output_path = rel_path
        exported_dest_by_doc_id[item.document_id] = dest
        stats["total_exported"] += 1
        stats[category] += 1
        if collision_detected:
            stats["collisions"] += 1
            item.observations.append(
                f"colision_nombre_pdf: ya existía otro archivo para {item.candidate_fiscal_key}; derivado a {rel_path}"
            )
            item.requires_human_review = True
        if note:
            item.observations.append(note)

        size_mb = dest.stat().st_size / (1024.0 * 1024.0)
        over_limit = size_mb > _SUNAT_FILE_LIMIT_MB
        if over_limit:
            stats["over_limit_mb"] += 1
            item.observations.append(f"archivo_pdf_supera_10mb ({size_mb:.2f} MB)")
            item.requires_human_review = True

        # Exportamos los anexos o soportes del mismo PDF junto al comprobante en la carpeta del proveedor.
        linked_labels: list[str] = []
        linked_items = supports_by_owner_doc_id.get(item.document_id, [])
        if linked_items and category in {"cdp_complete", "notes_complete"}:
            support_ranges = [(sup.page_start, sup.page_end) for sup in linked_items]
            min_p = min(p1 for p1, _ in support_ranges)
            max_p = max(p2 for _, p2 in support_ranges)
            sup_dest = dest.with_name(f"{dest.stem}_SOPORTES_p{min_p}-{max_p}.pdf")
            try:
                _write_pdf_pages_list(unit, support_ranges, sup_dest)
                sup_rel = sup_dest.relative_to(output_root).as_posix()
                sup_size_mb = sup_dest.stat().st_size / (1024.0 * 1024.0)
                if sup_size_mb > _SUNAT_FILE_LIMIT_MB:
                    stats["over_limit_mb"] += 1
                    item.observations.append(f"archivo_soportes_supera_10mb ({sup_size_mb:.2f} MB)")
                for sup in linked_items:
                    sup.normalized_filename = sup_dest.name
                    sup.output_path = sup_rel
                    sup.observations.append(f"vinculado_a_cdp_mismo_pdf: {rel_path}")
                    sup_kind = sup.non_receipt_type or "soporte"
                    linked_labels.append(f"{sup_kind} (p.{sup.page_start}-{sup.page_end}) -> {sup_dest.name}")
                stats["linked_supports"] += len(linked_items)
                stats["total_exported"] += 1
            except Exception as exc:
                item.observations.append(f"error_exportacion_soportes_vinculados: {exc}")
                stats["errors"] += 1

        _set_pdf_derived_info(
            item,
            size_mb=size_mb,
            pages=exported_pages,
            over_limit=over_limit,
            linked_supports=linked_labels,
        )

    return stats
