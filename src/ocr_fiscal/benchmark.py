from __future__ import annotations

from collections import Counter
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tempfile import TemporaryDirectory
import time

from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from zipfile import BadZipFile

from .amounts import AMOUNT_COLUMNS, AMOUNT_FIELDS
from .config import load_config
from .exports import DATE_FORMAT, _style_and_populate_sheet, save_workbook
from .models import AMBIGUOUS, DETECTED, NOT_RECEIPT, READ_ERROR
from .processor import FiscalProcessor
from .tax import normalize_correlativo
from .visible import visible_labels


# Las cinco columnas que identifican un segmento en la referencia: archivo, páginas y bloques.
LOCATOR_COLUMNS = ["Archivo", "Página inicial", "Página final", "Bloque inicial", "Bloque final"]
# Los campos que se comparan, cada uno con el título de su columna en la plantilla y su nombre interno.
FIELDS = [
    ("Detección", "detection_status"),
    ("RUC emisor", "emitter_ruc"),
    ("RUC receptor", "recipient_ruc"),
    ("Tipo", "document_type"),
    ("Serie", "series"),
    ("Correlativo", "correlativo_normalized"),
    ("Fecha", "issue_date"),
    ("Moneda", "currency"),
    ("IGV", "igv"),
    ("Total", "total"),
    ("Descripción", "general_description"),
    ("Detracción visible", "detraccion_visible"),
    ("Retención visible", "retencion_visible"),
    ("Buen contribuyente visible", "buen_contribuyente_visible"),
]
# Columnas de la plantilla antigua, que todavía se aceptan al leer referencias.
LEGACY_REFERENCE_COLUMNS = LOCATOR_COLUMNS + [label for label, _ in FIELDS]
# A los campos anteriores se les suman los importes.
FIELDS += [(label, field) for label, field in AMOUNT_COLUMNS if field not in {"igv", "total"}]
REFERENCE_COLUMNS = LOCATOR_COLUMNS + [label for label, _ in FIELDS]
# Columnas que se guardan como texto en Excel para no perder los ceros iniciales (por ejemplo, el tipo 01).
TEXT_LABELS = {"RUC emisor", "RUC receptor", "Tipo", "Serie", "Correlativo"}
# Valores válidos de la columna «Detección» de la referencia.
DETECTION_STATES = {DETECTED, NOT_RECEIPT, AMBIGUOUS, READ_ERROR}


# Exige que la ruta termine en .xlsx.
def _xlsx_path(path: Path) -> None:
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"Use una ruta de Excel con extensión .xlsx: {path}")


# Comando plantilla-calibracion: crea el Excel vacío donde se anota, mirando los originales, lo que debería leerse.
# No sobrescribe una plantilla que ya exista.
def write_reference_template(path: Path) -> Path:
    _xlsx_path(path)
    if path.exists():
        raise ValueError(f"La plantilla ya existe: {path}")
    book = Workbook()
    sheet = book.active
    sheet.title = "Referencia"
    _style_and_populate_sheet(sheet, [(label, label) for label in REFERENCE_COLUMNS], [], set())
    # Excel debe conservar el cero inicial de tipos como 01 y 03.
    for index, label in enumerate(REFERENCE_COLUMNS, start=1):
        if label in TEXT_LABELS:
            sheet.column_dimensions[sheet.cell(1, index).column_letter].number_format = "@"
        elif label == "Fecha":
            sheet.column_dimensions[sheet.cell(1, index).column_letter].number_format = DATE_FORMAT
    return save_workbook(book, path)


# Lee la referencia anotada y la valida: hoja «Referencia», columnas de la plantilla, ubicaciones sin repetir
# y estados de detección válidos. Devuelve lo esperado por ubicación (archivo, páginas y bloques).
def _read_reference(path: Path) -> dict:
    _xlsx_path(path)
    try:
        book = load_workbook(path, read_only=True, data_only=True)
    except (InvalidFileException, BadZipFile) as exc:
        raise ValueError(f"No se pudo abrir la referencia de Excel {path}.") from exc
    expected = {}
    try:
        if "Referencia" not in book.sheetnames:
            raise ValueError("La referencia debe tener una hoja llamada Referencia.")
        rows = book["Referencia"].values
        headers = list(next(rows, ()))
        if headers not in (REFERENCE_COLUMNS, LEGACY_REFERENCE_COLUMNS):
            raise ValueError("Use las columnas de plantilla-calibracion en el orden original.")
        for index, row in enumerate(rows, start=2):
            if all(value is None for value in row):
                continue
            file, *coordinates = row[:5]
            if not isinstance(file, str) or not file.strip() or any(type(n) is not int or n < 1 for n in coordinates):
                raise ValueError(f"Revise el archivo, las páginas y los bloques de la fila {index}.")
            if coordinates[0] > coordinates[1] or coordinates[2] > coordinates[3]:
                raise ValueError(f"Los rangos de página o bloque están invertidos en la fila {index}.")
            locator = (file.strip().replace("\\", "/"), *coordinates)
            if locator in expected or row[5] not in DETECTION_STATES:
                raise ValueError(f"Ubicación duplicada o detección desconocida en la fila {index}.")
            fields = row[5:]
            expected[locator] = fields + (None,) * (len(FIELDS) - len(fields))
    finally:
        book.close()
    if not expected:
        raise ValueError("La referencia debe contener por lo menos una fila etiquetada.")
    return expected


# Deja un valor en una forma que se pueda comparar sin que importe cómo lo guardó Excel:
# fechas en ISO, correlativos de 8 dígitos, importes como Decimal y espacios sobrantes fuera.
def _comparable(value, field: str):
    if value is None:
        return None
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    if field == "issue_date" and isinstance(value, str):
        try:
            return datetime.strptime(value.strip(), "%d/%m/%Y").date().isoformat()
        except ValueError:
            pass
    if field == "correlativo_normalized":
        return normalize_correlativo(str(value)).normalized or str(value)
    if field in AMOUNT_FIELDS:
        try:
            return Decimal(str(value))
        except InvalidOperation:
            pass
    return " ".join(str(value).split())


# Comando calibrar: lee la misma carpeta con uno o varios perfiles de ajustes y compara, campo por campo,
# contra la referencia revisada. Devuelve un Excel con un resumen por perfil y otra hoja con el detalle.
def run_calibration(input_root: Path, reference: Path, output: Path, recipient: str, profiles: list[Path]) -> Path:
    _xlsx_path(output)
    input_root, reference, output = (path.resolve() for path in (input_root, reference, output))
    if not input_root.is_dir():
        raise ValueError(f"No existe la carpeta de muestra {input_root}.")
    if output == reference or input_root in output.parents or input_root in reference.parents:
        raise ValueError("Guarde la referencia y el informe fuera de la carpeta de muestra, en archivos distintos.")
    # Paso 1: leer la referencia.
    expected = _read_reference(reference)
    metrics, detail = [], []
    # Paso 2: por cada perfil, correr un análisis completo en una carpeta temporal y medir cuánto tarda.
    for number, profile in enumerate(profiles, start=1):
        if not profile.is_file():
            raise ValueError(f"No existe el perfil {profile}.")
        config = load_config(profile)
        config.recipient_ruc = recipient
        label = f"{number}. {profile.name}"
        with TemporaryDirectory(prefix="ocr-cdp-calibracion-") as temporary:
            processor = FiscalProcessor(Path(temporary) / "estado", config)
            start = time.perf_counter()
            summary = processor.scan(input_root)
            elapsed = round(time.perf_counter() - start, 3)
        # Paso 3: emparejar cada segmento leído con su fila de la referencia (por archivo, páginas y bloques)
        # y comparar campo por campo. Un segmento sin fila de referencia se anota como «Sin pareja».
        counts = Counter()
        for item in summary.outcomes:
            filename = Path(item.original_path).relative_to(input_root).as_posix()
            if item.member_path:
                filename += "!" + "!".join([*item.container_chain[1:], item.member_path])
            locator = (filename, item.page_start, item.page_end, item.block_start, item.block_end)
            actual = {**item.fields, "detection_status": item.detection_status, "general_description": item.general_description}
            for field in ("detraccion_visible", "retencion_visible", "buen_contribuyente_visible"):
                actual[field] = visible_labels(item.field_evidence, field)
            counts[locator] += 1
            reference_fields = expected.get(locator)
            if reference_fields is None or counts[locator] > 1:
                detail.append([label, *locator, "Segmento", None, item.detection_status, "Sin pareja en referencia"])
                continue
            for (title, field), value in zip(FIELDS, reference_fields, strict=True):
                if value is None:
                    continue
                observed = actual.get(field)
                status = "Coincide" if _comparable(value, field) == _comparable(observed, field) else "Difiere"
                detail.append([label, *locator, title, value, observed, status])
        # Lo que estaba en la referencia y no apareció en la lectura se anota como «No encontrado».
        for locator in expected.keys() - counts.keys():
            detail.append([label, *locator, "Segmento", expected[locator][0], None, "No encontrado"])
        # Paso 4: totales del perfil: cuántos campos coinciden y cuántos difieren, cuántos recuadros de importes
        # se aceptaron y cuántas alertas de OCR dudoso hubo.
        profile_detail = [row for row in detail if row[0] == label]
        evaluated_fields = Counter(row[-1] for row in profile_detail)
        metrics.append([
            label, config.render_dpi, config.languages, round(config.min_ocr_confidence * 100), config.timeout_seconds,
            summary.processed_file_count, len(expected), len(expected.keys() - counts.keys()),
            evaluated_fields["Sin pareja en referencia"], evaluated_fields["Coincide"],
            evaluated_fields["Difiere"],
            sum(
                check.get("reading_source") == "Recuadro de importes"
                for outcome in summary.outcomes for check in outcome.derived_checks
            ),
            sum(
                check.get("check") == "low_ocr_confidence"
                for outcome in summary.outcomes for check in outcome.derived_checks
            ),
            elapsed,
        ])

    # Paso 5: escribir el Excel con dos hojas: «Perfiles» (un resumen por perfil) y «Detalle» (cada comparación).
    book = Workbook()
    book.active.title = "Perfiles"
    headers = [
        "Archivo de ajustes", "Resolución PDF", "Idiomas", "Confianza mínima de 0 a 100", "Límite OCR (s)", "Archivos procesados",
        "Segmentos de referencia", "Segmentos faltantes", "Segmentos sin pareja", "Campos que coinciden",
        "Campos que difieren", "Recuadros aceptados", "Alertas de lectura", "Duración (s)",
    ]
    _style_and_populate_sheet(book.active, [(key, key) for key in headers], [dict(zip(headers, row)) for row in metrics], set())
    headers = ["Perfil", *LOCATOR_COLUMNS, "Campo", "Esperado", "Leído", "Resultado"]
    _style_and_populate_sheet(
        book.create_sheet("Detalle"), [(key, key) for key in headers],
        [dict(zip(headers, row)) for row in detail], set(),
    )
    return save_workbook(book, output)
