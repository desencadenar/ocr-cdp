"""Segunda lectura opcional del cuadro de importes de una factura escaneada."""

from dataclasses import replace
import math
from pathlib import Path
import re
from tempfile import TemporaryDirectory

from PIL import Image, ImageDraw, ImageOps
import pymupdf

from .amounts import SETTLEMENT_FIELDS, amount_mentions, reconcile_amounts
from .archives import WorkUnit
from .config import AppConfig
from .documents import PageRead
from .extraction import _record_amount_field
from .models import ExtractionResult, LayoutLine, StageEvent
from .ocr import OCREngine, OCRExecutionError, OCRUnavailable, TesseractOCREngine


# Rótulos que marcan dónde empieza y dónde acaba el cuadro de importes:
# «Sub total ventas» arriba e «Importe total» abajo.
_SUBTOTAL = re.compile(r"^\s*SUB\s*TOTAL\s+VENTAS?\b", re.IGNORECASE)
_TOTAL = re.compile(r"^\s*IMPORTE\s+TOTAL\b", re.IGNORECASE)
# Importes que el cuadro tiene que traer por sí mismo para que se acepte la segunda lectura.
_REQUIRED_FIELDS = {"subtotal", "valor_venta", "descuentos", "igv", "total"}
# Resolución a la que se recorta y se vuelve a leer el cuadro (más alta que la de la página entera).
_PANEL_DPI = 300
# Margen blanco, en píxeles, que se añade alrededor del recorte para que el OCR no lea mal los bordes.
_BORDER = 12


# Busca en el OCR de la página dónde empieza y dónde acaba el cuadro de importes y devuelve ese rectángulo.
# Si no hay exactamente un inicio y un final, o su forma no tiene sentido, devuelve None.
def _panel_box(layout: list[LayoutLine]) -> pymupdf.Rect | None:
    # Paso 1: localizar la línea que abre el cuadro y la que lo cierra. Tiene que haber exactamente una de cada.
    starts = [line.bbox for line in layout if line.bbox and line.source == "tesseract" and _SUBTOTAL.search(line.text)]
    ends = [line.bbox for line in layout if line.bbox and line.source == "tesseract" and _TOTAL.search(line.text)]
    if len(starts) != 1 or len(ends) != 1:
        return None
    # Paso 2: comprobar que la forma es razonable: el final debajo del inicio y ambos más o menos alineados.
    start, end = starts[0], ends[0]
    height = max(start[3] - start[1], end[3] - end[1])
    if height <= 0 or end[1] <= start[3] or abs(start[2] - end[2]) > 2 * height:
        return None
    # Paso 3: el rectángulo, con un pequeño margen para no cortar las cifras.
    return pymupdf.Rect(
        min(start[0], end[0]) - max(16, height * 0.65),
        start[1] - height * 0.3,
        max(start[2], end[2]) + max(15, height * 0.5),
        end[3] + height * 0.3,
    )


# Quita las líneas de la tabla (rayas largas, horizontales o verticales) de la copia temporal del recorte,
# porque el OCR las confunde con caracteres.
def _without_grid(image: Image.Image) -> Image.Image:
    # Solo se borran trazos largos y rectos en la copia temporal. Los trazos
    # de letras y cifras son demasiado cortos para cumplir este criterio.
    # Pasamos a gris y marcamos como «oscuro» lo que queda por debajo de un umbral.
    gray = ImageOps.grayscale(image)
    dark = gray.point(lambda pixel: pixel < 135, "1").load()
    width, height = gray.size
    draw = ImageDraw.Draw(gray)
    # Recorremos primero las filas y luego las columnas buscando rayas continuas.
    for vertical in (False, True):
        span, tracks = (height, width) if vertical else (width, height)
        for track in range(tracks):
            start = None
            for offset in range(span + 1):
                filled = offset < span and dark[(track, offset) if vertical else (offset, track)]
                if filled and start is None:
                    start = offset
                elif not filled and start is not None:
                    # Una raya que cubre al menos el 40 % del ancho (o del alto) del recorte es parte de la tabla: la pintamos de blanco.
                    if offset - start >= span * 0.4:
                        bounds = (
                            (track - 1, start, track + 1, offset - 1)
                            if vertical else (start, track - 1, offset - 1, track + 1)
                        )
                        draw.rectangle(bounds, fill=255)
                    start = None
    # Devolvemos la imagen con el margen blanco alrededor.
    return ImageOps.expand(gray, _BORDER, fill=255)


# Recorta el cuadro de importes del PDF original, lo vuelve a leer con Tesseract en modo de bloque único (psm 6)
# y devuelve sus líneas con las coordenadas ya convertidas a las de la página completa.
def _read_region(
    unit: WorkUnit, page: PageRead, box: pymupdf.Rect,
    config: AppConfig, engine: TesseractOCREngine,
) -> list[LayoutLine]:
    # Paso 1: abrir el PDF, desde memoria si viene de dentro de un ZIP o desde el disco si no.
    document = (
        pymupdf.open(stream=unit.content_bytes, filetype="pdf")
        if unit.content_bytes is not None else pymupdf.open(unit.disk_path)
    )
    with document:
        original = document[page.page_number - 1]
        # Paso 2: pasar el rectángulo de la resolución de página a puntos del PDF.
        # Si queda vacío, es demasiado alto (más del 40 % de la página) o pasa del tamaño máximo de imagen, no se lee.
        clip = (box * (72 / config.render_dpi)) & original.rect
        scale = _PANEL_DPI / 72
        if clip.is_empty or clip.height > original.rect.height * 0.4:
            return []
        if (math.ceil(clip.width * scale) + 2 * _BORDER) * (math.ceil(clip.height * scale) + 2 * _BORDER) > config.max_pixels:
            return []
        # Paso 3: renderizar sólo ese recorte a 300 DPI, quitarle la cuadrícula y reconocerlo con Tesseract.
        pixmap = original.get_pixmap(dpi=_PANEL_DPI, clip=clip, alpha=False)
        image = _without_grid(Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples))
        with TemporaryDirectory(prefix="cdp-importes-") as directory:
            path = Path(directory) / "recuadro.png"
            image.save(path, dpi=(_PANEL_DPI, _PANEL_DPI))
            recognized = engine.recognize(path, page_segmentation_mode=6)

    # Paso 4: devolver las líneas con sus coordenadas en la escala de la página completa,
    # marcadas como lectura del recuadro.
    ratio = config.render_dpi / _PANEL_DPI
    lines = []
    for line in recognized.layout_lines:
        if not line.bbox:
            continue
        x0, y0, x1, y1 = line.bbox
        bbox = (
            (pixmap.x + x0 - _BORDER) * ratio, (pixmap.y + y0 - _BORDER) * ratio,
            (pixmap.x + x1 - _BORDER) * ratio, (pixmap.y + y1 - _BORDER) * ratio,
        )
        lines.append(replace(line, bbox=bbox, source="ocr_recuadro_importes"))
    return lines


# Decide si la segunda lectura del cuadro se acepta y, si se acepta, reemplaza los importes de la página.
# Sólo se acepta cuando el cuadro trae por sí mismo todos los importes necesarios,
# su total coincide con el de la página y la aritmética cuadra.
def apply_amount_panel(result: ExtractionResult, layout: list[LayoutLine]) -> bool:
    """Acepta la lectura solo si el cuadro tiene sus propios importes y cuadra."""
    # Paso 1: extraer los importes del recuadro en un resultado aparte, sin tocar todavía el original.
    # Se ignoran los importes dudosos y los de detracción, retención, percepción, anticipos y total a pagar.
    panel = ExtractionResult(currency=result.currency)
    lines = [line.text for line in layout]
    for index, line in enumerate(lines):
        for field, observed, currency, dudoso in amount_mentions(line):
            if dudoso or field in SETTLEMENT_FIELDS or field in {"anticipos", "total_a_pagar"}:
                continue
            _record_amount_field(
                panel, field, observed, index, lines, result.page_start, result.block_start,
                {}, layout, currency, source="ocr_recuadro_importes",
            )
    # Paso 2: condiciones para aceptar. El total del recuadro es el mismo que el de la página,
    # trae todos los importes requeridos y la aritmética cuadra.
    if not result.total or panel.total != result.total:
        return False
    if any(getattr(panel, field) is None for field in _REQUIRED_FIELDS):
        return False
    if reconcile_amounts(panel)["status"] != "cuadra":
        return False

    # Paso 3: probamos primero sobre una copia que sustituye los importes por los del recuadro,
    # conservando las lecturas anteriores marcadas como reemplazadas. Si aun así no cuadra, se descarta.
    fields = {field for field in panel.field_evidence if getattr(panel, field) is not None}
    combined = replace(result, candidates=dict(result.candidates), field_evidence=dict(result.field_evidence))
    for field in fields:
        setattr(combined, field, getattr(panel, field))
        combined.candidates.pop(field, None)
        combined.field_evidence[field] = [
            replace(original, field_status="lectura_pagina_reemplazada")
            for original in combined.field_evidence.get(field, [])
        ] + panel.field_evidence[field]
    revised = reconcile_amounts(combined)
    if revised["status"] != "cuadra":
        return False
    # Dejamos constancia de que este cuadre salió del recuadro.
    revised["reading_source"] = "Recuadro de importes"

    # Paso 4: aceptado. Se retira el aviso anterior de que no cuadraba, para que no queden mensajes contradictorios.
    before = next(
        (check for check in result.derived_checks if check["check"] == "arithmetic_total_reconciliation"), None
    )
    if before:
        previous_reason = before.get("reason") or (
            f"Los componentes suman {before.get('expected_total')} y el total leído es {before.get('observed_total')}."
        )
        if previous_reason in result.observations:
            result.observations.remove(previous_reason)
    # Paso 5: pasamos al resultado real los importes del recuadro, con su evidencia y su confianza.
    for field in fields:
        setattr(result, field, getattr(panel, field))
        result.candidates.pop(field, None)
        result.fields_original[field] = panel.fields_original[field]
        result.field_statuses[field] = panel.field_statuses[field]
        result.field_evidence[field] = combined.field_evidence[field]
        result.evidence.setdefault(field, []).extend(panel.evidence[field])
        result.confidence[field] = panel.confidence[field]
    # Sustituimos el control aritmético anterior por el nuevo, que sí cuadra.
    result.derived_checks = [check for check in result.derived_checks if check is not before] + [revised]
    if "amounts" in result.secondary_review_fields:
        result.secondary_review_fields.remove("amounts")
    result.observations.append("Importes cotejados con una lectura separada del recuadro y el mismo total.")
    return True


# Punto de entrada de este módulo: decide si merece la pena releer el cuadro de esta página
# y devuelve un evento con el resultado (o None si no hacía falta).
def review_amount_panel(
    unit: WorkUnit, page: PageRead, result: ExtractionResult,
    config: AppConfig, engine: OCREngine,
) -> StageEvent | None:
    # Sólo se relee en páginas leídas por OCR con Tesseract que tengan un total.
    if not isinstance(engine, TesseractOCREngine) or page.source != "ocr" or not result.total:
        return None
    # Y sólo si la aritmética todavía no cuadra: si ya cuadra, no hay nada que mejorar.
    if any(check["check"] == "arithmetic_total_reconciliation" and check["status"] == "cuadra" for check in result.derived_checks):
        return None
    # Si no se encuentra el cuadro, no se hace nada.
    box = _panel_box(page.layout_lines)
    if box is None:
        return None
    try:
        accepted = apply_amount_panel(result, _read_region(unit, page, box, config, engine))
    except (OSError, RuntimeError, OCRExecutionError, OCRUnavailable) as exc:
        return StageEvent("ocr_recuadro_importes", "failed", str(exc), page.page_number)
    # El resultado queda en el registro de eventos de la página.
    message = "El recuadro cuadra con el total leído en la página." if accepted else (
        "El recuadro no aportó importes completos que cuadren con el total de la página."
    )
    return StageEvent("ocr_recuadro_importes", "succeeded" if accepted else "not_accepted", message, page.page_number)
