from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import tempfile
import time
from typing import Any
import xml.etree.ElementTree as ET

import pymupdf
from PIL import Image, ImageStat

from .archives import WorkUnit, detect_format_from_bytes
from .config import AppConfig
from .models import LayoutLine, StageEvent
from .ocr import OCRExecutionError, OCREngine, OCRResult, OCRUnavailable, TesseractOCREngine


# DEPRECADO: sólo lo usa read_document(), que nadie llama.
SUPPORTED_SUFFIXES = {
    ".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp", ".heic", ".xml",
}


# Se lanza cuando un documento no se puede abrir o leer (archivo dañado, formato no admitido...).
class DocumentReadError(RuntimeError):
    pass


# Lo que se leyó de una página: el texto, de dónde salió (texto nativo del PDF u OCR), su estado,
# unas medidas de calidad y las líneas con su posición.
@dataclass
class PageRead:
    page_number: int
    text: str
    source: str
    status: str
    quality: dict[str, int | float | str]
    ocr_confidence: float | None = None
    line_confidences: dict[str, float] = field(default_factory=dict)
    layout_lines: list[LayoutLine] = field(default_factory=list)
    structured_data: dict[str, Any] | None = None
    error_message: str | None = None


# Lo que se leyó de un documento completo: sus páginas, los eventos del proceso y el motor de OCR que se usó.
@dataclass
class DocumentRead:
    path: Path
    source_format: str
    page_count: int
    pages: list[PageRead]
    events: list[StageEvent] = field(default_factory=list)
    status: str = "read"
    ocr_engine: str = "none"
    ocr_version: str = "none"

    # Todo el texto del documento, una página tras otra.
    @property
    def text(self) -> str:
        return "\n\n".join(page.text for page in self.pages if page.text)


# Medidas básicas de una imagen (ancho, alto y variación de grises). Se guardan para saber con qué calidad se leyó.
def _quality(width: int, height: int, grayscale_stddev: float) -> dict[str, int | float]:
    return {"width": width, "height": height, "grayscale_stddev": round(grayscale_stddev, 2)}


# Lo mismo, calculado a partir de una imagen.
def _quality_from_image(image: Image.Image) -> dict[str, int | float]:
    grayscale_stddev = ImageStat.Stat(image.convert("L")).stddev[0]
    return _quality(image.width, image.height, grayscale_stddev)


# Eventos fijos que se anotan en cada página para dejar claro qué mejoras de imagen NO se aplican
# en esta versión (girar, enderezar, rectificar, subir contraste, segunda lectura).
def _static_events(page_number: int) -> list[StageEvent]:
    return [
        StageEvent("rotation_correction", "not_applied", "V1 registra orientación pero no rota automáticamente", page_number),
        StageEvent("deskew", "not_applied", "corrección de inclinación pendiente para una versión posterior", page_number),
        StageEvent("perspective_rectification", "not_applied", "rectificación geométrica no incluida en V1", page_number),
        StageEvent("contrast_enhancement", "not_applied", "mejora selectiva no aplicada en V1", page_number),
        StageEvent("alternate_region_ocr", "not_applied", "no se ejecuta un segundo motor/región en V1", page_number),
    ]


# Comprueba que el motor de OCR está disponible y copia al registro de eventos lo que averiguó al comprobarlo.
def _check_engine(engine: OCREngine | None, events: list[StageEvent] | None = None) -> tuple[bool, str | None]:
    if engine is None:
        return False, "No hay un motor OCR configurado"
    try:
        prev_probes = len(getattr(engine, "probe_events", []) or [])
        available, reason = engine.availability()
        if events is not None:
            probe_events = list(getattr(engine, "probe_events", []) or [])
            for probe_ev in probe_events[prev_probes:]:
                probe_kind = str(probe_ev.get("probe") or "probe")
                events.append(
                    StageEvent(
                        stage=f"ocr_{probe_kind}_probe",
                        status=str(probe_ev.get("status") or ("succeeded" if available else "failed")),
                        message=f"Sonda Tesseract ({probe_kind})",
                        details=dict(probe_ev),
                    )
                )
        return available, reason
    except Exception as exc:
        return False, f"Error comprobando el motor OCR: {exc}"


# Pasa una imagen por el OCR y devuelve: texto, motor usado, estado, confianza y líneas con su posición.
# Si el OCR no está o falla, lo anota y devuelve un resultado vacío con el motivo.
def _ocr_image(
    image_path: Path,
    page_number: int,
    engine: OCREngine | None,
    available: bool,
    unavailable_reason: str | None,
    events: list[StageEvent],
) -> tuple[str, str, str, float | None, dict[str, float], list[LayoutLine], str | None]:
    started = time.perf_counter()
    if not available or engine is None:
        err_msg = unavailable_reason or "Motor OCR no disponible"
        events.append(
            StageEvent(
                "ocr",
                "failed",
                err_msg,
                page_number,
                duration_ms=(time.perf_counter() - started) * 1000.0,
            )
        )
        return "", "", "ocr_unavailable", None, {}, [], err_msg
    try:
        # La lectura de verdad. Si falla, se devuelve vacío y con el error.
        result: OCRResult = engine.recognize(image_path)
    except (OCRExecutionError, OCRUnavailable, OSError) as exc:
        err_msg = str(exc)
        events.append(
            StageEvent(
                "ocr",
                "failed",
                err_msg,
                page_number,
                duration_ms=(time.perf_counter() - started) * 1000.0,
            )
        )
        return "", getattr(engine, "name", "tesseract"), "ocr_error", None, {}, [], err_msg
    duration_ms = (time.perf_counter() - started) * 1000.0
    events.append(
        StageEvent(
            "ocr",
            "succeeded" if result.text.strip() else "empty",
            "Tesseract local",
            page_number,
            duration_ms=duration_ms,
        )
    )
    # Si el motor no entregó líneas con posición, las armamos a partir del texto (sin coordenadas).
    layout_lines = list(getattr(result, "layout_lines", []) or [])
    if not layout_lines and result.text.strip():
        for idx, line in enumerate(result.text.strip().splitlines(), start=1):
            if line.strip():
                layout_lines.append(
                    LayoutLine(
                        line_number=idx,
                        text=line.strip(),
                        block_number=1,
                        bbox=None,
                        ocr_confidence=result.line_confidences.get(line.strip(), result.mean_confidence),
                        source="tesseract",
                    )
                )
    return (
        result.text.strip(),
        getattr(engine, "name", "tesseract"),
        "read" if result.text.strip() else "empty_text",
        result.mean_confidence,
        dict(result.line_confidences),
        layout_lines,
        None,
    )


# Saca el texto nativo de una página de PDF junto con la posición de cada línea.
# Las palabras se agrupan en líneas por su altura, con una tolerancia de unos 3 puntos.
def _extract_native_pdf_layout(page: pymupdf.Page, y_tol: float = 3.0) -> tuple[str, list[LayoutLine]]:
    raw_text = page.get_text("text").strip()
    if not raw_text:
        return "", []
    words = page.get_text("words")
    if not words:
        lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
        layout = [
            LayoutLine(line_number=i, text=line, block_number=1, bbox=None, ocr_confidence=None, source="native_pdf_text")
            for i, line in enumerate(lines, start=1)
        ]
        return "\n".join(lines), layout

    # Paso 1: ordenar las palabras de arriba abajo y de izquierda a derecha, y agruparlas por filas.
    words_sorted = sorted(words, key=lambda w: (round(float(w[1]) / y_tol), float(w[0])))
    grouped_rows: list[list[Any]] = []
    cur_row: list[Any] = []
    cur_y: float | None = None
    for word in words_sorted:
        wy = float(word[1])
        if cur_y is None or abs(wy - cur_y) <= y_tol:
            cur_row.append(word)
            if cur_y is None:
                cur_y = wy
        else:
            cur_row.sort(key=lambda item: float(item[0]))
            grouped_rows.append(cur_row)
            cur_row = [word]
            cur_y = wy
    if cur_row:
        cur_row.sort(key=lambda item: float(item[0]))
        grouped_rows.append(cur_row)

    # Paso 2: de cada fila armamos una línea de texto, con su caja contenedora y el bloque al que pertenece.
    layout_lines: list[LayoutLine] = []
    text_lines: list[str] = []
    for idx, row in enumerate(grouped_rows, start=1):
        line_text = " ".join(str(item[4]).strip() for item in row if str(item[4]).strip()).strip()
        if not line_text:
            continue
        bbox = (
            round(min(float(item[0]) for item in row), 2),
            round(min(float(item[1]) for item in row), 2),
            round(max(float(item[2]) for item in row), 2),
            round(max(float(item[3]) for item in row), 2),
        )
        block_no = int(row[0][5]) + 1 if len(row[0]) > 5 else 1
        layout_lines.append(
            LayoutLine(
                line_number=idx,
                text=line_text,
                block_number=max(1, block_no),
                bbox=bbox,
                ocr_confidence=None,
                source="native_pdf_text",
            )
        )
        text_lines.append(line_text)
    return "\n".join(text_lines), layout_lines


# DEPRECADO: sólo lo usa el lector de XML, que ya no se ejecuta.
def _local_tag(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


# DEPRECADO: lector de facturas XML (UBL). Los XML se excluyen desde el inventario y nunca llegan aquí.
# Se conserva por si algún día se retoma ese flujo; el resto de este archivo no depende de él.
def _parse_xml_bytes(raw_bytes: bytes, source_path: Path) -> DocumentRead:
    started = time.perf_counter()
    events: list[StageEvent] = [StageEvent("format_detection", "succeeded", "xml", 1)]
    text_decoded = raw_bytes.decode("utf-8", errors="replace")
    if "<!ENTITY" in text_decoded.upper() or "<!DOCTYPE" in text_decoded.upper():
        raise DocumentReadError("XML con declaraciones DOCTYPE/ENTITY bloqueado por seguridad local")
    try:
        root = ET.fromstring(raw_bytes)
    except ET.ParseError as exc:
        raise DocumentReadError(f"No se pudo analizar XML: {exc}") from exc

    root_name = _local_tag(root.tag)
    type_map = {
        "Invoice": ("01", "FACTURA ELECTRÓNICA"),
        "CreditNote": ("07", "NOTA DE CRÉDITO ELECTRÓNICA"),
        "DebitNote": ("08", "NOTA DE DÉBITO ELECTRÓNICA"),
    }
    doc_code, doc_title = type_map.get(root_name, (None, root_name))

    notes: dict[str, str] = {}
    top_id: str | None = None
    issue_date: str | None = None
    due_date: str | None = None
    currency: str | None = None
    type_code_obs: str | None = None
    supplier_ruc: str | None = None
    supplier_name: str | None = None
    customer_ruc: str | None = None
    customer_name: str | None = None
    payment_condition: str | None = None
    tax_amount: str | None = None
    taxable_amount: str | None = None
    inafecto_amount: str | None = None
    exonerado_amount: str | None = None
    line_extension_amount: str | None = None
    payable_amount: str | None = None
    referenced_doc: str | None = None
    items: list[dict[str, str | None]] = []

    for child in root:
        cname = _local_tag(child.tag)
        val = (child.text or "").strip()
        if cname == "ID" and val and top_id is None:
            top_id = val
        elif cname == "IssueDate" and val:
            issue_date = val
        elif cname == "DueDate" and val:
            due_date = val
        elif cname in {"InvoiceTypeCode", "CreditNoteTypeCode", "DebitNoteTypeCode"} and val:
            type_code_obs = val
        elif cname == "DocumentCurrencyCode" and val:
            currency = val
        elif cname == "Note" and val:
            loc = child.attrib.get("languageLocaleID", "note")
            notes[loc] = val
        elif cname == "UBLExtensions":
            for prop in child.iter():
                if _local_tag(prop.tag) == "AdditionalProperty":
                    pid = ""
                    pval = ""
                    for sub in prop:
                        sname = _local_tag(sub.tag)
                        if sname == "ID":
                            pid = (sub.text or "").strip()
                        elif sname == "Value":
                            pval = (sub.text or "").strip()
                    if pid and pval:
                        notes[pid] = pval
        elif cname == "BillingReference":
            for sub in child.iter():
                if _local_tag(sub.tag) == "ID" and (sub.text or "").strip():
                    referenced_doc = (sub.text or "").strip()
                    break
        elif cname == "AccountingSupplierParty":
            for sub in child.iter():
                sname = _local_tag(sub.tag)
                sval = (sub.text or "").strip()
                if sname in {"ID", "CustomerAssignedAccountID"} and len(sval) == 11 and sval.isdigit():
                    supplier_ruc = sval
                elif sname == "RegistrationName" and sval:
                    supplier_name = sval
                elif sname == "Name" and sval and not supplier_name:
                    supplier_name = sval
        elif cname == "AccountingCustomerParty":
            for sub in child.iter():
                sname = _local_tag(sub.tag)
                sval = (sub.text or "").strip()
                if sname in {"ID", "CustomerAssignedAccountID"} and len(sval) == 11 and sval.isdigit():
                    customer_ruc = sval
                elif sname == "RegistrationName" and sval:
                    customer_name = sval
                elif sname == "Name" and sval and not customer_name:
                    customer_name = sval
        elif cname == "PaymentTerms":
            p_id = ""
            p_means = ""
            for sub in child:
                sname = _local_tag(sub.tag)
                sval = (sub.text or "").strip()
                if sname == "ID":
                    p_id = sval
                elif sname == "PaymentMeansID":
                    p_means = sval
                elif sname == "PaymentDueDate" and sval and not due_date:
                    due_date = sval
            if p_id == "FormaPago" and p_means and not payment_condition:
                payment_condition = p_means
        elif cname == "TaxTotal":
            for sub in child:
                sname = _local_tag(sub.tag)
                sval = (sub.text or "").strip()
                if sname == "TaxAmount" and sval and tax_amount is None:
                    tax_amount = sval
                elif sname == "TaxSubtotal":
                    sub_taxable = None
                    sub_tax = None
                    sub_scheme = ""
                    for tnode in sub.iter():
                        tname = _local_tag(tnode.tag)
                        tval = (tnode.text or "").strip()
                        if tname == "TaxableAmount" and tval:
                            sub_taxable = tval
                        elif tname == "TaxAmount" and tval:
                            sub_tax = tval
                        elif tname == "Name" and tval:
                            sub_scheme = tval.upper()
                    if sub_scheme in {"IGV", "VAT"} and sub_taxable:
                        taxable_amount = sub_taxable
                        if sub_tax:
                            tax_amount = sub_tax
                    elif sub_scheme in {"INA", "INAFECTO"} and sub_taxable:
                        inafecto_amount = sub_taxable
                    elif sub_scheme in {"EXO", "EXONERADO"} and sub_taxable:
                        exonerado_amount = sub_taxable
        elif cname == "LegalMonetaryTotal":
            for sub in child:
                sname = _local_tag(sub.tag)
                sval = (sub.text or "").strip()
                if sname == "LineExtensionAmount" and sval:
                    line_extension_amount = sval
                elif sname == "PayableAmount" and sval:
                    payable_amount = sval
                elif sname == "TaxInclusiveAmount" and sval and not payable_amount:
                    payable_amount = sval
        elif cname in {"InvoiceLine", "CreditNoteLine", "DebitNoteLine"}:
            item_id = None
            qty = None
            unit = None
            line_ext = None
            unit_price = None
            desc_parts: list[str] = []
            code = None
            for sub in child:
                sname = _local_tag(sub.tag)
                sval = (sub.text or "").strip()
                if sname == "ID" and sval:
                    item_id = sval
                elif sname in {"InvoicedQuantity", "CreditedQuantity", "DebitedQuantity"} and sval:
                    qty = sval
                    unit = sub.attrib.get("unitCode")
                elif sname == "LineExtensionAmount" and sval:
                    line_ext = sval
                elif sname == "PricingReference":
                    for pnode in sub.iter():
                        if _local_tag(pnode.tag) == "PriceAmount" and (pnode.text or "").strip():
                            unit_price = (pnode.text or "").strip()
                elif sname == "Price" and not unit_price:
                    for pnode in sub.iter():
                        if _local_tag(pnode.tag) == "PriceAmount" and (pnode.text or "").strip():
                            unit_price = (pnode.text or "").strip()
                elif sname == "Item":
                    for inode in sub:
                        iname = _local_tag(inode.tag)
                        ival = (inode.text or "").strip()
                        if iname == "Description" and ival:
                            desc_parts.append(ival)
                        elif iname == "SellersItemIdentification":
                            for cnode in inode.iter():
                                if _local_tag(cnode.tag) == "ID" and (cnode.text or "").strip():
                                    code = (cnode.text or "").strip()
            items.append(
                {
                    "item_id": item_id,
                    "code": code,
                    "description": " - ".join(desc_parts) if desc_parts else None,
                    "quantity": qty,
                    "unit": unit,
                    "unit_price": unit_price,
                    "line_total": line_ext,
                }
            )

    structured = {
        "root_tag": root_name,
        "document_type_code": type_code_obs if type_code_obs in {"01", "03", "07", "08"} else doc_code,
        "document_title": doc_title,
        "id": top_id,
        "issue_date": issue_date,
        "due_date": due_date,
        "currency": currency,
        "supplier_ruc": supplier_ruc,
        "supplier_name": supplier_name,
        "customer_ruc": customer_ruc,
        "customer_name": customer_name,
        "payment_condition": payment_condition,
        "referenced_document": referenced_doc,
        "subtotal": taxable_amount or line_extension_amount,
        "inafecto": inafecto_amount,
        "exonerado": exonerado_amount,
        "igv": tax_amount,
        "total": payable_amount,
        "general_description": notes.get("SUMMARY") or notes.get("SUSTENTO") or notes.get("note"),
        "notes": notes,
        "items": items,
    }

    synth_lines: list[str] = []
    if doc_title:
        synth_lines.append(doc_title)
    if top_id:
        synth_lines.append(f"ID UBL: {top_id}")
    if supplier_name:
        synth_lines.append(f"RAZÓN SOCIAL DEL EMISOR: {supplier_name}")
    if supplier_ruc:
        synth_lines.append(f"RUC EMISOR: {supplier_ruc}")
    if customer_name:
        synth_lines.append(f"RAZÓN SOCIAL ADQUIRENTE: {customer_name}")
    if customer_ruc:
        synth_lines.append(f"RUC ADQUIRENTE: {customer_ruc}")
    if issue_date:
        synth_lines.append(f"FECHA DE EMISIÓN: {issue_date}")
    if due_date:
        synth_lines.append(f"FECHA DE VENCIMIENTO: {due_date}")
    if currency:
        synth_lines.append(f"MONEDA: {currency}")
    if payment_condition:
        synth_lines.append(f"FORMA DE PAGO: {payment_condition}")
    if structured["general_description"]:
        synth_lines.append(f"GLOSA: {structured['general_description']}")
    for idx, item in enumerate(items, start=1):
        synth_lines.append(
            f"ITEM {idx}: {item.get('description') or ''} | CANT: {item.get('quantity') or ''} "
            f"{item.get('unit') or ''} | P.UNIT: {item.get('unit_price') or ''} | VALOR VENTA: {item.get('line_total') or ''}"
        )
    if structured["subtotal"]:
        synth_lines.append(f"VALOR DE VENTA: {structured['subtotal']}")
    if inafecto_amount:
        synth_lines.append(f"INAFECTO: {inafecto_amount}")
    if tax_amount is not None:
        synth_lines.append(f"IGV: {tax_amount}")
    if payable_amount:
        synth_lines.append(f"TOTAL: {payable_amount}")

    layout_lines = [
        LayoutLine(line_number=i, text=line, block_number=1, bbox=None, ocr_confidence=None, source="xml_structured")
        for i, line in enumerate(synth_lines, start=1)
    ]
    text = "\n".join(synth_lines)
    duration_ms = (time.perf_counter() - started) * 1000.0
    events.append(StageEvent("xml_structured_read", "succeeded", f"UBL {root_name} ({len(items)} ítems)", 1, duration_ms=duration_ms))
    page = PageRead(
        page_number=1,
        text=text,
        source="xml_structured",
        status="read" if text else "empty",
        quality={"xml_bytes": len(raw_bytes), "line_count": len(synth_lines)},
        ocr_confidence=None,
        line_confidences={},
        layout_lines=layout_lines,
        structured_data=structured,
    )
    return DocumentRead(
        path=source_path,
        source_format="xml",
        page_count=1,
        pages=[page],
        events=events,
        status="read" if text else "empty",
        ocr_engine="none",
        ocr_version="none",
    )


# Decide, para cada página de PDF, si alcanza con su texto nativo o si hay que hacerle OCR.
def _evaluate_pdf_page_reading_mode(
    page: pymupdf.Page,
    native_text: str,
    layout_lines: list[LayoutLine],
) -> tuple[str, str, dict[str, int | float | str]]:
    """Decide si el texto nativo basta o si la página requiere OCR sobre contenido rasterizado.

    Un pie de página seleccionable, un sello o una marca de archivo pueden acompañar
    una factura escaneada. Por ello medimos qué proporción de la página ocupan las
    imágenes rasterizadas y qué cobertura vertical/textual aporta la capa nativa.
    """
    # Paso 1: medir la página y cuánto de ella ocupan las imágenes.
    page_w = max(float(page.rect.width), 1.0)
    page_h = max(float(page.rect.height), 1.0)
    page_area = page_w * page_h

    image_count = 0
    total_img_area = 0.0
    max_img_area = 0.0
    document_images: list[pymupdf.Rect] = []
    try:
        img_infos = page.get_image_info()
        image_count = len(img_infos)
        for info in img_infos:
            bbox = info.get("bbox")
            if bbox and len(bbox) == 4:
                x0 = max(float(bbox[0]), float(page.rect.x0))
                y0 = max(float(bbox[1]), float(page.rect.y0))
                x1 = min(float(bbox[2]), float(page.rect.x1))
                y1 = min(float(bbox[3]), float(page.rect.y1))
                if x1 > x0 and y1 > y0:
                    area = (x1 - x0) * (y1 - y0)
                    total_img_area += area
                    max_img_area = max(max_img_area, area)
                    if area >= 0.15 * page_area:
                        document_images.append(pymupdf.Rect(x0, y0, x1, y1))
    except Exception:
        image_count = len(page.get_images(full=True))

    image_coverage_ratio = round(min(total_img_area / page_area, 1.0), 4)
    max_image_ratio = round(min(max_img_area / page_area, 1.0), 4)

    # Paso 2: medir el texto nativo: cuántos caracteres y líneas tiene y qué parte de la altura de la página cubre.
    stripped_native = native_text.strip()
    native_chars = len(stripped_native)
    native_lines = len(layout_lines)

    bboxes = [line.bbox for line in layout_lines if line.bbox is not None]
    if bboxes:
        min_y = min(b[1] for b in bboxes)
        max_y = max(b[3] for b in bboxes)
        vertical_span_ratio = round(max(0.0, (max_y - min_y) / page_h), 4)
        confined_to_margin = (max_y <= 0.22 * page_h) or (min_y >= 0.72 * page_h)
    else:
        vertical_span_ratio = 0.0
        confined_to_margin = True

    metrics: dict[str, int | float | str] = {
        "native_text_chars": native_chars,
        "layout_lines": native_lines,
        "image_count": image_count,
        "image_coverage_ratio": image_coverage_ratio,
        "max_image_ratio": max_image_ratio,
        "native_vertical_span_ratio": vertical_span_ratio,
    }

    # Paso 3: reglas de decisión. Sin texto nativo, OCR.
    if not stripped_native:
        metrics["reading_decision"] = "ocr_pagina_sin_texto_nativo"
        return "ocr", "sin texto nativo seleccionable", metrics

    # Si no hay una imagen grande que ocupe la página, el texto nativo basta.
    has_document_raster = image_coverage_ratio >= 0.15 or max_image_ratio >= 0.15
    if not has_document_raster:
        metrics["reading_decision"] = "nativo_sin_imagen_documental"
        return "native_pdf_text", "texto PDF nativo sin imagen rasterizada dominante", metrics

    # Cuando coexisten una imagen documental grande y poco texto nativo (o confinado al margen),
    # el texto nativo suele ser un pie, encabezado o sello sobre un documento escaneado.
    is_marginal_native = (
        confined_to_margin
        or native_lines < 4
        or native_chars < 140
        or vertical_span_ratio < 0.18
    )
    if is_marginal_native:
        metrics["reading_decision"] = "ocr_sobre_pagina_hibrida_con_texto_marginal"
        return (
            "ocr",
            f"imagen rasterizada ({image_coverage_ratio:.0%} área) con texto nativo marginal ({native_chars} caracteres)",
            metrics,
        )

    # El texto de otra zona de la página no da cobertura a una imagen documental.
    for image in document_images:
        inside = [bbox for bbox in bboxes if (pymupdf.Rect(bbox) & image).get_area() >= 0.5 * pymupdf.Rect(bbox).get_area()]
        vertical_coverage = max((b[3] for b in inside), default=0) - min((b[1] for b in inside), default=0)
        if len(inside) < 4 or vertical_coverage < 0.18 * image.height:
            metrics["reading_decision"] = "ocr_imagen_sin_cobertura_nativa"
            return "ocr", "imagen documental sin texto nativo representativo en su zona", metrics

    # Si ninguna imagen grande se queda sin texto que la cubra, el texto nativo es representativo.
    metrics["reading_decision"] = "nativo_representativo"
    return "native_pdf_text", "texto PDF nativo representativo con coordenadas", metrics


# Resultado explícito para los XML: se dejan fuera y se anota el motivo.
def _excluded_xml_document_read(source_path: Path) -> DocumentRead:
    """Devuelve un resultado explícito de exclusión por formato para archivos XML."""
    reason = "Archivo XML excluido por formato del flujo visual de comprobantes (solo PDF e imágenes)"
    events = [
        StageEvent("format_detection", "excluded_by_format", "xml"),
        StageEvent("format_exclusion", "excluded_by_format", reason, 1),
    ]
    page = PageRead(
        page_number=1,
        text="",
        source="excluded_xml",
        status="excluded_by_format",
        quality={"excluded_xml": 1},
        ocr_confidence=None,
        line_confidences={},
        layout_lines=[],
        structured_data=None,
    )
    return DocumentRead(
        path=source_path,
        source_format="xml",
        page_count=1,
        pages=[page],
        events=events,
        status="excluded_by_format",
        ocr_engine="none",
        ocr_version="none",
    )


# Lee un PDF página por página. Para cada una decide entre texto nativo y OCR;
# si necesita OCR, la convierte en imagen a la resolución configurada y la pasa por Tesseract.
def _read_pdf_document(
    source_path: Path,
    raw_bytes: bytes | None,
    engine: OCREngine | None,
    cfg: AppConfig,
) -> DocumentRead:
    events: list[StageEvent] = [StageEvent("format_detection", "succeeded", "pdf")]
    engine_state: tuple[bool, str | None] | None = None
    active_engine = engine

    # Comprueba el motor de OCR sólo la primera vez que hace falta (y lo crea si no nos lo dieron).
    def get_engine_state() -> tuple[bool, str | None]:
        nonlocal engine_state, active_engine
        if active_engine is None:
            active_engine = TesseractOCREngine(cfg.tesseract_path, cfg.languages, cfg.timeout_seconds)
        if engine_state is None:
            engine_state = _check_engine(active_engine, events)
        return engine_state

    pages: list[PageRead] = []
    used_engine = "none"
    used_version = "none"

    try:
        # Paso 1: abrir el PDF, desde memoria si venía de un ZIP o desde el disco.
        pdf = pymupdf.open(stream=raw_bytes, filetype="pdf") if raw_bytes is not None else pymupdf.open(source_path)
    except Exception as exc:
        raise DocumentReadError(f"No se pudo abrir PDF: {exc}") from exc

    try:
        events.append(StageEvent("page_count", "succeeded", str(pdf.page_count)))
        # Paso 2: recorrer las páginas una por una.
        for index, page in enumerate(pdf, start=1):
            page_started = time.perf_counter()
            page_events = _static_events(index)
            # Primero el texto nativo, con su posición.
            native_text, layout_lines = _extract_native_pdf_layout(page)
            page_events.append(
                StageEvent(
                    "native_text",
                    "succeeded" if native_text else "empty",
                    f"{len(native_text)} caracteres",
                    index,
                    duration_ms=(time.perf_counter() - page_started) * 1000.0,
                )
            )
            # Luego se decide si ese texto basta.
            reading_mode, reading_reason, decision_metrics = _evaluate_pdf_page_reading_mode(
                page, native_text, layout_lines
            )
            # Caso A: el texto nativo basta, no se hace OCR.
            if reading_mode == "native_pdf_text":
                page_events.append(StageEvent("quality_analysis", "succeeded", reading_reason, index))
                page_events.append(StageEvent("ocr", "skipped", "se usó texto nativo representativo", index))
                pages.append(
                    PageRead(
                        page_number=index,
                        text=native_text,
                        source="native_pdf_text",
                        status="read",
                        quality=decision_metrics,
                        layout_lines=layout_lines,
                    )
                )
                events.extend(page_events)
                continue

            # Caso B: hace falta OCR. Primero se comprueba que la página no pase del tamaño máximo de imagen;
            # si pasa, queda sin leer y se anota el motivo.
            scale = cfg.render_dpi / 72.0
            estimated_pixels = int(page.rect.width * scale) * int(page.rect.height * scale)
            page_events.append(
                StageEvent("quality_analysis", "succeeded", f"{reading_reason}; estimado {estimated_pixels} píxeles", index)
            )
            if estimated_pixels > cfg.max_pixels:
                reason = f"página supera límite {cfg.max_pixels} píxeles"
                page_events.append(StageEvent("render", "failed", reason, index))
                pages.append(
                    PageRead(
                        index,
                        "",
                        "none",
                        "quality_limit",
                        {**decision_metrics, "estimated_pixels": estimated_pixels},
                        error_message=reason,
                    )
                )
                events.extend(page_events)
                continue
            # Se convierte la página en imagen a la resolución configurada.
            render_started = time.perf_counter()
            pixmap = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
            quality_image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
            quality: dict[str, int | float | str] = {**decision_metrics, **_quality_from_image(quality_image)}
            with tempfile.TemporaryDirectory(prefix="ocr-fiscal-") as temp_dir:
                image_path = Path(temp_dir) / f"page-{index:04d}.png"
                # Guardar el PNG con metadato DPI explícito evita que Tesseract asuma ~96 DPI
                # y pierda líneas de serie/correlativo en tickets térmicos rasterizados (p. ej. F401-00000123).
                if getattr(cfg, "explicit_render_dpi_metadata", True):
                    quality_image.save(image_path, format="PNG", dpi=(cfg.render_dpi, cfg.render_dpi))
                else:
                    pixmap.save(image_path)
                page_events.append(
                    StageEvent(
                        "render",
                        "succeeded",
                        f"{pixmap.width}x{pixmap.height}@{cfg.render_dpi}dpi",
                        index,
                        duration_ms=(time.perf_counter() - render_started) * 1000.0,
                    )
                )
                # Se comprueba Tesseract y se lee la imagen de la página.
                available, unavailable_reason = get_engine_state()
                text, engine_name, page_status, confidence, line_confidences, ocr_layout, err_msg = _ocr_image(
                    image_path, index, active_engine, available, unavailable_reason, page_events
                )
            used_engine = engine_name if engine_name != "" else used_engine
            if active_engine is not None and engine_name != "":
                used_version = active_engine.version

            # Si la página híbrida tenía un pie/sello nativo que el OCR no capturó y el OCR tuvo éxito,
            # conservamos la evidencia del pie al final sin reemplazar el cuerpo reconocido por OCR.
            combined_text = text
            combined_layout = list(ocr_layout)
            if text and native_text.strip():
                existing_upper = text.upper()
                extra_native = [
                    ln for ln in layout_lines if ln.text.strip() and ln.text.strip().upper() not in existing_upper
                ]
                if extra_native:
                    extra_str = "\n".join(ln.text.strip() for ln in extra_native)
                    combined_text = f"{text}\n{extra_str}".strip()
                    base_line_no = len(combined_layout)
                    for offset, ln in enumerate(extra_native, start=1):
                        combined_layout.append(
                            LayoutLine(
                                line_number=base_line_no + offset,
                                text=ln.text.strip(),
                                block_number=ln.block_number,
                                bbox=ln.bbox,
                                ocr_confidence=ln.ocr_confidence,
                                source="native_pdf_text",
                            )
                        )

            pages.append(
                PageRead(
                    page_number=index,
                    text=combined_text,
                    source="ocr" if text else "none",
                    status=page_status,
                    quality=quality,
                    ocr_confidence=confidence,
                    line_confidences=line_confidences,
                    layout_lines=combined_layout,
                    error_message=err_msg,
                )
            )
            events.extend(page_events)
    finally:
        pdf.close()

    # Resumen del documento según lo que se pudo leer de sus páginas.
    return _finalize_document_read(source_path, "pdf", pages, events, used_engine, used_version)


# Lee una imagen (o un TIFF de varias páginas) con OCR, página por página, siempre que no supere el tamaño máximo.
def _read_image_document(
    source_path: Path,
    raw_bytes: bytes | None,
    fmt: str,
    engine: OCREngine | None,
    cfg: AppConfig,
) -> DocumentRead:
    import io

    events: list[StageEvent] = [StageEvent("format_detection", "succeeded", fmt)]
    engine_state: tuple[bool, str | None] | None = None
    active_engine = engine

    def get_engine_state() -> tuple[bool, str | None]:
        nonlocal engine_state, active_engine
        if active_engine is None:
            active_engine = TesseractOCREngine(cfg.tesseract_path, cfg.languages, cfg.timeout_seconds)
        if engine_state is None:
            engine_state = _check_engine(active_engine, events)
        return engine_state

    pages: list[PageRead] = []
    used_engine = "none"
    used_version = "none"

    try:
        # Paso 1: abrir la imagen.
        image = Image.open(io.BytesIO(raw_bytes)) if raw_bytes is not None else Image.open(source_path)
    except Exception as exc:
        raise DocumentReadError(f"No se pudo abrir imagen: {exc}") from exc

    try:
        # Paso 2: una imagen puede tener varios cuadros (un TIFF multipágina). Cada cuadro cuenta como una página.
        frame_count = getattr(image, "n_frames", 1)
        events.append(StageEvent("page_count", "succeeded", str(frame_count)))
        for index in range(frame_count):
            image.seek(index)
            page_number = index + 1
            page_events = _static_events(page_number)
            width, height = image.size
            pixels = width * height
            # Si es demasiado grande, no se lee y se anota el motivo.
            if pixels > cfg.max_pixels:
                reason = f"imagen supera límite {cfg.max_pixels} píxeles"
                page_events.append(StageEvent("quality_analysis", "failed", reason, page_number))
                page_events.append(StageEvent("ocr", "skipped", "no se convirtió la imagen", page_number))
                pages.append(
                    PageRead(
                        page_number,
                        "",
                        "none",
                        "quality_limit",
                        {"width": width, "height": height},
                        error_message=reason,
                    )
                )
                events.extend(page_events)
                continue
            # Se guarda el cuadro como PNG temporal y se lee con OCR.
            frame = image.convert("RGB")
            quality = _quality_from_image(frame)
            page_events.append(StageEvent("quality_analysis", "succeeded", f"{pixels} píxeles", page_number))
            with tempfile.TemporaryDirectory(prefix="ocr-fiscal-") as temp_dir:
                image_path = Path(temp_dir) / f"frame-{page_number:04d}.png"
                frame.save(image_path, format="PNG")
                available, unavailable_reason = get_engine_state()
                text, engine_name, page_status, confidence, line_confidences, ocr_layout, err_msg = _ocr_image(
                    image_path, page_number, active_engine, available, unavailable_reason, page_events
                )
            used_engine = engine_name if engine_name != "" else used_engine
            if active_engine is not None and engine_name != "":
                used_version = active_engine.version
            pages.append(
                PageRead(
                    page_number=page_number,
                    text=text,
                    source="ocr" if text else "none",
                    status=page_status,
                    quality=quality,
                    ocr_confidence=confidence,
                    line_confidences=line_confidences,
                    layout_lines=ocr_layout,
                    error_message=err_msg,
                )
            )
            events.extend(page_events)
    finally:
        image.close()

    return _finalize_document_read(source_path, fmt, pages, events, used_engine, used_version)


# Decide el estado global del documento según sus páginas: leído, parcial (algunas sin texto),
# sin OCR disponible, con error de OCR o vacío.
def _finalize_document_read(
    source_path: Path,
    fmt: str,
    pages: list[PageRead],
    events: list[StageEvent],
    used_engine: str,
    used_version: str,
) -> DocumentRead:
    if not pages:
        status = "empty"
    elif all(page.text for page in pages):
        status = "read"
    elif any(page.text for page in pages):
        status = "partial"
    elif any(page.status == "ocr_unavailable" for page in pages):
        status = "ocr_unavailable"
    elif any(page.status == "ocr_error" for page in pages):
        status = "ocr_error"
    else:
        status = "empty"
    return DocumentRead(source_path, fmt, len(pages), pages, events, status, used_engine, used_version)


# Punto de entrada de este módulo: lee una unidad de trabajo según su formato (PDF o imagen)
# y devuelve el DocumentRead. Los XML se excluyen aquí también.
def read_work_unit(unit: WorkUnit, engine: OCREngine | None, config: AppConfig | None = None) -> DocumentRead:
    cfg = config or AppConfig()
    if unit.error:
        raise DocumentReadError(unit.error)
    fmt = unit.source_ref.detected_format.lower()
    source_path = unit.disk_path or Path(unit.source_ref.source_path)
    # Si el archivo viene del disco, se comprueba que siga existiendo.
    raw_bytes = unit.content_bytes
    if raw_bytes is None and unit.disk_path is not None:
        if not unit.disk_path.is_file():
            raise DocumentReadError(f"No existe el archivo: {unit.disk_path}")

    # Según el formato: XML excluido, PDF o imagen. Cualquier otro formato es un error.
    if fmt == "xml":
        return _excluded_xml_document_read(source_path)
    if fmt == "pdf":
        return _read_pdf_document(source_path, raw_bytes, engine, cfg)
    if fmt in {"jpg", "jpeg", "png", "tif", "tiff", "bmp", "webp", "heic"}:
        return _read_image_document(source_path, raw_bytes, fmt, engine, cfg)
    raise DocumentReadError(f"Formato no admitido: {fmt}")


# DEPRECADO: nadie lo llama. Hacía lo mismo que read_work_unit, pero partiendo de una ruta en disco.
def read_document(path: str | Path, engine: OCREngine | None = None, config: AppConfig | None = None) -> DocumentRead:
    source = Path(path)
    cfg = config or AppConfig()
    suffix = source.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise DocumentReadError(f"Formato no admitido: {suffix or '(sin extensión)'}")
    if not source.is_file():
        raise DocumentReadError(f"No existe el archivo: {source}")

    with source.open("rb") as handle:
        header = handle.read(64)
    detected_fmt = detect_format_from_bytes(header, source.name)

    if suffix == ".xml" or detected_fmt == "xml":
        return _excluded_xml_document_read(source)
    if suffix == ".pdf" or detected_fmt == "pdf":
        return _read_pdf_document(source, None, engine, cfg)
    return _read_image_document(source, None, suffix[1:], engine, cfg)
