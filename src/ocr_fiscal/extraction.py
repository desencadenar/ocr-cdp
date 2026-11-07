from __future__ import annotations

from decimal import Decimal, InvalidOperation
import re
from datetime import datetime
from typing import Any, Mapping, Sequence

from .amounts import AMOUNT_FIELDS, amount_mentions, reconcile_amounts
from .models import (
    AMBIGUOUS,
    DEFAULT_DOCUMENT_TYPES,
    DETECTED,
    IDENTITY_COMPLETE,
    IDENTITY_CONFLICT,
    IDENTITY_INCOMPLETE,
    IDENTITY_NOT_APPLICABLE,
    NOT_RECEIPT,
    READ_ERROR,
    ExtractionResult,
    FieldEvidence,
    LayoutLine,
    LineItem,
)
from .tax import normalize_correlativo, validate_ruc
from .visible import find_visible_mentions


# ===== Expresiones regulares =====
# Lo que sigue son los «moldes» con los que se reconoce cada pieza de un comprobante en el texto
# (RUC, serie, fechas, roles, títulos...). Se aplican línea por línea y casi todas ignoran mayúsculas y minúsculas.
#
# El rótulo «RUC» y sus variantes («R.U.C.»).
_RUC_LABEL = re.compile(r"\bR\s*[.:]?\s*U\s*[.:]?\s*C\s*[.:]?\b", re.IGNORECASE)
# Un posible número de RUC: 11 caracteres alfanuméricos aislados. Más adelante se validan.
_RUC_TOKEN = re.compile(r"(?<![A-Z0-9])([A-Z0-9]{11})(?![A-Z0-9])", re.IGNORECASE)

# Serie fiscal válida: empieza con F, B o E seguida de 2 o 3 caracteres alfanuméricos (con al menos un dígito),
# o son 3 o 4 dígitos cuando el contexto es claramente el de un comprobante.
# Acepta también formas como «FO02 N° - 00000123».
_SERIAL_PAIR = re.compile(
    r"(?<![A-Z0-9-])(?P<series>[FBE][A-Z0-9]{2,3}|[0-9]{3,4})\s*(?:N(?:RO|°|º|\.)\s*)?[-–]\s*(?P<number>[0-9]{1,20})(?![A-Z0-9-])",
    re.IGNORECASE,
)
# Serie escrita con su rótulo («Serie: F001»).
_SERIES_LABEL = re.compile(r"\bSERIE\s*[:#-]?\s*([A-Z0-9]{1,20})\b", re.IGNORECASE)
# Número escrito con su rótulo («Correlativo:», «N°», «Nro.»).
_NUMBER_LABEL = re.compile(
    r"\b(?:CORRELATIVO|N(?:RO|[ÚU]MERO|°|º))\s*[:#-]?\s*([0-9]{1,20})\b",
    re.IGNORECASE,
)
# Una fecha suelta, en año-mes-día o en día-mes-año.
_DATE = re.compile(
    r"\b(?:[0-9]{4}[-/][0-9]{1,2}[-/][0-9]{1,2}|[0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{4})\b"
)
# Entre el rótulo y la fecha el OCR suele dejar basura («€”», «+», rayas). Admitimos un
# tramo corto sin letras ni dígitos para no perder la fecha que sí está escrita en la página.
# Al excluir las letras, «FECHA DE VENCIMIENTO» y «FECHA DE VCTO» siguen sin tomarse por
# fecha de emisión.
_LABEL_NOISE = r"[^0-9A-Za-zÁÉÍÓÚÜÑáéíóúüñ]{0,8}"
_DATE_LABEL = re.compile(
    r"\bFECHA(?:\s+DE)?(?:\s+EMISI[ÓO]N)?" + _LABEL_NOISE
    + r"([0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{4}|[0-9]{4}-[0-9]{2}-[0-9]{2})",
    re.IGNORECASE,
)
# La fecha de vencimiento, con su rótulo.
_DUE_DATE_LABEL = re.compile(
    r"\bFECHA\s*DE\s*VENCIMIENTO\s*[:#-]?\s*([0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{4}|[0-9]{4}-[0-9]{2}-[0-9]{2})",
    re.IGNORECASE,
)
# Meses en español, para las fechas escritas con palabras («14 de Enero del 2019»).
_SPANISH_MONTHS: dict[str, int] = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}
# Una fecha escrita con palabras.
_SPANISH_TEXT_DATE = re.compile(
    r"\b([0-9]{1,2})\s+de\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\s+(?:del?|de\s+el)\s+([0-9]{4})\b",
    re.IGNORECASE,
)

# Cómo se escribe el título de cada tipo de comprobante: código SUNAT, nombre y expresión que lo detecta.
_TYPE_PATTERNS: tuple[tuple[str, str, re.Pattern[str]], ...] = (
    ("07", "Nota de crédito", re.compile(r"\bNOTA\s+(?:DE\s+)?CR[EÉ]DITO\b", re.IGNORECASE)),
    ("08", "Nota de débito", re.compile(r"\bNOTA\s+(?:DE\s+)?D[EÉ]BITO\b", re.IGNORECASE)),
    ("03", "Boleta de venta", re.compile(r"\bBOLETA(?:\s+DE\s+VENTA)?\b(?!\s+DE\s+(?!VENTA\b))", re.IGNORECASE)),
    ("01", "Factura", re.compile(r"\bFACTURA\b(?!\s+(?:PROFORMA|NEGOCIABLE)\b)", re.IGNORECASE)),
    ("02", "Recibo por honorarios", re.compile(r"\bRECIBO\s+(?:ELECTR[ÓO]NICO\s+)?POR\s+HONORARIOS\b", re.IGNORECASE)),
    ("04", "Liquidación de compra", re.compile(r"\bLIQUIDACI[ÓO]N\s+DE\s+COMPRA\b", re.IGNORECASE)),
    ("05", "Boleto de transporte aéreo", re.compile(r"\bBOLETO\s+DE\s+TRANSPORTE\s+A[EÉ]REO\b", re.IGNORECASE)),
    ("12", "Ticket de máquina registradora", re.compile(r"\b(?:TICKET|CINTA)\s+DE\s+M[ÁA]QUINA\s+REGISTRADORA\b", re.IGNORECASE)),
)
# Lo que puede ir delante del título en el mismo renglón: RUC, teléfono o el nombre del emisor con su forma societaria.
# En las páginas de dos columnas el OCR suele juntar renglones distintos.
_HEADER_CONTEXT = re.compile(
    r"\b(?:R\.?\s*U\.?\s*C|TEL(?:[ÉE]FONO|EF)?\.?)\s*[:.]?\s*[0-9]|"
    r"\bS\s*\.\s*(?:A(?:\s*\.\s*[AC])?|R\s*\.\s*L)\s*\.", re.IGNORECASE,
)
# «FACTURA ELECTRÓNICA» es una sola denominación; el adjetivo acompaña al título.
_TITLE_TAIL = re.compile(r"^\s*(ELECTR[ÓO]NIC[AO]|ELECTR?\.)(?![A-ZÁÉÍÓÚÑ])", re.IGNORECASE)
# Lo que puede ir detrás del título: «ELECTRÓNICA» y una serie con su correlativo.
_HEADER_SUFFIX = re.compile(
    r"(?:\s*(?:ELECTR[ÓO]NIC[AO]|ELECTR?\.?))?"
    r"(?:\s*(?:N(?:RO\.?|[°º.]|[ÚU]MERO)?\s*[:.#]?\s*)?(?:[A-Z0-9]{1,6}\s*[-–]\s*)?\d{1,20})?"
    r"\s*[:.]*\s*$",
    re.IGNORECASE,
)
# «Comprobante de compra» de pasajes aéreos: no tiene una denominación fiscal admitida y se manda a revisión.
_AIR_PURCHASE = re.compile(r"\bCOMPROBANTE\s+DE\s+COMPRA\b", re.IGNORECASE)
_AIR_CONTEXT = re.compile(r"\b(?:ITINERARIO\s+DE\s+VUELO|TARJETA\s+DE\s+EMBARQUE|N[ÚU]MERO\s+DE\s+TICKET)\b", re.IGNORECASE)

# Líneas donde «FACTURA» o «NOTA DE CRÉDITO» se mencionan como referencia, pie de página o anexo de factura negociable,
# y no como título del documento.
_NON_HEADER_TYPE_CONTEXT = re.compile(
    r"(?:"
    r"REPRESENTACI[ÓO]N\s+IMPRESA\s+DE\s+LA|"
    r"TIPO\s+DE\s+DOCUMENTO\s+QUE\s+MODIFICA|"
    r"TIPO\s+DE\s+NOTA\s+DE\s+CR[EÉ]DITO|"
    r"COD\.?\s+TIPO\s+NOTA|"
    r"DOCUMENTO\s+DE\s+REFERENCIA|"
    r"\b(?:REFERENCIA|REF\.?|SEG[ÚU]N)\s+(?:DE\s+)?(?:FACTURA|COMPROBANTE|DOCUMENTO|BOLETA)\b|"
    r"FACTURA\s+AFECTAD[AO]|"
    r"FACTURA\s+ELECT\.?\s+[FBE0-9]|"
    r"MODIFICACI[ÓO]N\s+DE\s+LA\s+FACTURA|"
    r"POR\s+LA\s+FACTURA|"
    r"\(\s*FACTURA\s+NRO|"
    r"TOTAL\s+PRESENTE\s+(?:FACTURA|NOTA)|"
    r"RECEPCI[ÓO]N\s+DE\s+ESTA\s+(?:FACTURA|NOTA)|"
    r"INFORMACI[ÓO]N\s+ADICIONAL\s+DE\s+LA\s+FACTURA\s+NEGOCIABLE|"
    r"FECHA\s+DE\s+ENTREGA\s+DE\s+LA\s+FACTURA\s+NEGOCIABLE|"
    r"EN\s+LA\s+FACTURA\s+NEGOCIABLE|"
    r"CL\.?\s*DOC\.?\s*:\s*RE\s*\(\s*FACTURA\s+BRUTO|"
    r"PARA\s+CONSULTAR\s+EL\s+COMPROBANTE"
    r")",
    re.IGNORECASE,
)

# Líneas donde una serie-número es la referencia a OTRO comprobante (por ejemplo, el que modifica una nota),
# no la identidad de este.
_REFERENCE_SERIAL_LINE = re.compile(
    r"(?:"
    r"\b(?:FACTURA|COMPROBANTE|DOCUMENTO)\s+(?:AFECTAD[AO]|DE\s+REFERENCIA|REFERENCIA|RELACIONAD[AO]|QUE\s+MODIFICA)\b|"
    r"\b(?:REFERENCIA|REF\.?|SEG[ÚU]N)\s+(?:DE\s+)?(?:FACTURA|COMPROBANTE|DOCUMENTO|BOLETA)\b|"
    r"SERIE\s+Y\s+N[ÚU]MERO\s+DEL\s+DOCUMENTO\s+QUE\s+MODIFICA|"
    r"TIPO\s+DE\s+NOTA\s+DE\s+(?:CR[EÉ]DITO|D[EÉ]BITO)|"
    r"POR\s+LOS\s+INTERESES\s+GENERADOS\s+POR\s+LA\s+FACTURA|"
    r"MODIFICACI[ÓO]N\s+DE\s+LA\s+FACTURA|"
    r"\(\s*FACTURA\s+NRO\.?|"
    r"FACTURA\s+ELECT\.?\s+[FBE][A-Z0-9]{2,3}"
    r")",
    re.IGNORECASE,
)

# Palabras que marcan quién es quién: emisor, receptor y terceros (transportistas, bancos, garantes...).
# Se usan para asignar el rol de cada RUC.
_ISSUER_ROLE = re.compile(r"\b(?:EMISORA?|ISSUING\s+AIRLINE|PROVEEDOR|VENDEDOR|EXPEDIDOR)\b", re.IGNORECASE)
_RECIPIENT_ROLE = re.compile(
    r"\b(?:ADQUIRENTE|ADQUIRIENTE|CLIENTE|RECEPTOR|USUARIO|SE[ÑN]ORE?\(?S?\)?|SEFIORE?\(?S?\)?|DESTINATARIO|COMPRADOR|FOID|D\.?\s*IDENTIDAD)\b",
    re.IGNORECASE,
)
_THIRD_PARTY_ROLE = re.compile(
    r"\b(?:TRANSPORTISTA|CONDUCTOR|TERCERO|GARANTE|CESIONARIO|FACTORING|BANCO\s+PAGADOR|AGENTE\s+DE\s+ADUANA)\b",
    re.IGNORECASE,
)
# Encabezados de documentos de soporte (estado de cuenta, HES, orden de compra, guía de remisión...).
# Sirven para separar un soporte de un comprobante cuando comparten página.
_SUPPORT_HEADER_LINE_RE = re.compile(
    r"(?:"
    r"ACCOUNT\s+STATEMENT\s+INQUIRY|"
    r"CONSULTA\s+DE\s+DECLARACI[ÓO]N\s+DE\s+CUENTA|"
    r"ESTADO\s+DE\s+CUENTA\s+(?:CORRIENTE|BANCARIO)|"
    r"HOJA\s+DE\s+ENTRADA\s+DE\s+SERVICIOS|"
    r"ACEPTACI[ÓO]N\s+DE\s+SERVICIOS|"
    r"N[°º]\s*HOJA\s+ENTRADA\s+SERV|"
    r"RUTA\s+DE\s+APROBACI[ÓO]N|"
    r"\bORDEN\s+DE\s+COMPRA\b|"
    r"\bPURCHASE\s+ORDER\b|"
    r"\bGU[ÍI]A\s+DE\s+REMISI[ÓO]N\b|"
    r"VISUALIZAR\s+DOCUMENTO\s*:\s*VISTA\s+DE\s+ENTRADA|"
    r"CL\.?\s*DOC\.?\s*:\s*RE\s*\(\s*FACTURA\s+BRUTO"
    r")",
    re.IGNORECASE,
)
# Señales de que una línea forma parte del membrete del emisor: domicilio fiscal, forma societaria, teléfono, calle o avenida.
_HEADER_COMPANY_CUE = re.compile(
    r"\b(?:DOMICILIO\s+FISCAL|S\.A\.A\.|S\.A\.C\.|S\.A\.|S\.R\.L\.|E\.I\.R\.L\.|TEL[EÉ]F(?:ONO)?|URB\.|PROLONGACI[ÓO]N|AV\.|AVENIDA|CALLE|JR\.|JIR[ÓO]N|OFICINA)\b",
    re.IGNORECASE,
)
# Formas societarias escritas con puntos (S.A.C., S.R.L., E.I.R.L....).
_LEGAL_ENTITY_SUFFIX = re.compile(
    r"\b(?:S\.A\.A\.|S\.A\.C\.|S\.A\.|S\.R\.L\.|E\.I\.R\.L\.)(?=\s|$)",
    re.IGNORECASE,
)
# Muchas empresas imprimen el tipo societario sin puntos («EMPRESA EJEMPLO SAC»). Lo
# aceptamos solo al final del renglón, donde cierra el nombre, para no confundir estas
# tres o cuatro letras con una palabra cualquiera del encabezado.
_LEGAL_ENTITY_SUFFIX_PLAIN = re.compile(r"\b(?:SAA|SAC|SRL|EIRL)\s*\.?\s*$", re.IGNORECASE)
# Textos genéricos como «comprobante electrónico de pago», que por sí solos no bastan para decidir el tipo.
_GENERIC_RECEIPT = re.compile(
    r"\b(?:COMPROBANTE\s+(?:ELECTR[ÓO]NICO\s+)?DE\s+PAGO|COMPROBANTE\s+ELECTR[ÓO]NICO)\b",
    re.IGNORECASE,
)
# «Página 2 de 3»: señal de que esta página continúa a la anterior.
_PAGE_MARKER = re.compile(r"\bP[ÁA]G(?:INA|\.)?\s*:?\s*[0-9]{1,3}\s+(?:DE|/)\s*[0-9]{1,3}\b", re.IGNORECASE)
# Palabras que indican continuación o anexo («continúa», «información adicional de la factura negociable»...).
_CONTINUATION = re.compile(
    r"\b(?:CONTIN[ÚU]A|CONTINUACI[ÓO]N|SIGUE|INFORMACI[ÓO]N\s+ADICIONAL\s+DE\s+LA\s+FACTURA\s+NEGOCIABLE|FECHA\s+DE\s+ENTREGA\s+DE\s+LA\s+FACTURA\s+NEGOCIABLE|ANEXO\s+DE\s+FACTURACI[ÓO]N)\b",
    re.IGNORECASE,
)
# Estados de página que cuentan como error de lectura.
_READ_ERRORS = {"ocr_error", "ocr_unavailable", "quality_limit", "read_error", "failed"}


# Minúsculas y espacios simples, para comparar líneas sin que importen las mayúsculas ni los espacios.
def _norm(line: str) -> str:
    return re.sub(r"\s+", " ", line).strip().casefold()


# Pasa un importe tal como se leyó («1,234.50», «34 926,71», «164.231,59») a un número con punto decimal y dos decimales.
# Si no se entiende, devuelve None.
def _normalize_amount_str(raw: str | None) -> str | None:
    if not raw:
        return None
    cleaned = re.sub(r"[^\d,.\s-]", "", raw).strip()
    if not cleaned:
        return None
    # Formato «34 926,71»: espacios para los miles y coma para los decimales.
    if re.fullmatch(r"-?\d{1,3}(?:\s+\d{3})+,\d{2}", cleaned):
        cleaned = cleaned.replace(" ", "").replace(",", ".")
    elif re.fullmatch(r"-?\d+,\d{2}", cleaned):
        cleaned = cleaned.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(?:,\d{3})+,\d{2}", cleaned):
        # Confusión típica del OCR: el último punto decimal se leyó como coma («164,231,59»).
        last_comma = cleaned.rfind(",")
        cleaned = cleaned[:last_comma].replace(",", "") + "." + cleaned[last_comma + 1 :]
    elif re.fullmatch(r"-?\d{1,3}(?:\.\d{3})+,\d{2}", cleaned):
        cleaned = cleaned.replace(".", "").replace(",", ".")
    else:
        cleaned = cleaned.replace(" ", "").replace(",", "")
    try:
        val = Decimal(cleaned)
        return f"{val:.2f}"
    except InvalidOperation:
        return None


def _normalize_quantity_str(raw: str | None) -> str | None:
    """Normaliza cantidades preservando hasta 3 decimales (p. ej. 1.000) sin aplicar reglas de moneda."""
    if not raw:
        return None
    cleaned = re.sub(r"[^\d,.-]", "", raw).strip()
    if not cleaned:
        return None
    if re.fullmatch(r"\d+[.,]\d{3}", cleaned):
        cleaned = cleaned.replace(",", ".")
    elif re.fullmatch(r"\d+,\d{1,2}", cleaned):
        cleaned = cleaned.replace(",", ".")
    else:
        cleaned = cleaned.replace(",", "")
    try:
        val = Decimal(cleaned)
        return f"{val:.2f}"
    except InvalidOperation:
        return None


# Confianza del OCR para un grupo de líneas: la de la línea más débil.
# Sale de la posición de cada línea o, si no la hay, de buscar su texto en el mapa de confianzas.
def _line_confidence(
    lines: list[str],
    indexes: Sequence[int],
    confidence_map: Mapping[str, float],
    layout_lines: Sequence[LayoutLine] | None = None,
) -> float | None:
    scores: list[float] = []
    if layout_lines:
        for idx in indexes:
            if 0 <= idx < len(layout_lines) and layout_lines[idx].ocr_confidence is not None:
                scores.append(float(layout_lines[idx].ocr_confidence))
    if not scores and confidence_map:
        by_normalized = {_norm(str(key)): float(value) for key, value in confidence_map.items()}
        for idx in indexes:
            if 0 <= idx < len(lines) and _norm(lines[idx]) in by_normalized:
                scores.append(by_normalized[_norm(lines[idx])])
    return min(scores) if scores else None


# Rectángulo que contiene todas las líneas indicadas (en coordenadas de la página). None si no hay posiciones.
def _line_bbox(
    indexes: Sequence[int],
    layout_lines: Sequence[LayoutLine] | None = None,
) -> tuple[float, float, float, float] | None:
    if not layout_lines:
        return None
    boxes = [
        layout_lines[idx].bbox
        for idx in indexes
        if 0 <= idx < len(layout_lines) and layout_lines[idx].bbox is not None
    ]
    if not boxes:
        return None
    return (
        round(min(b[0] for b in boxes if b is not None), 2),
        round(min(b[1] for b in boxes if b is not None), 2),
        round(max(b[2] for b in boxes if b is not None), 2),
        round(max(b[3] for b in boxes if b is not None), 2),
    )


# Fragmento de texto que sirve de apoyo: las líneas indicadas más algunas de contexto por arriba y por abajo.
def _excerpt(lines: list[str], indexes: Sequence[int], radius: int = 1) -> str:
    valid = [index for index in indexes if 0 <= index < len(lines)]
    if not valid:
        return ""
    start = max(0, min(valid) - radius)
    end = min(len(lines), max(valid) + radius + 1)
    return "\n".join(line.strip() for line in lines[start:end] if line.strip())


# Registra un dato leído junto con su evidencia: valor observado y normalizado, página, fragmento, confianza,
# posición y estado. Todo dato que acaba en el Excel pasa por aquí.
def _add_field(
    result: ExtractionResult,
    field_name: str,
    observed: str,
    normalized: str | None,
    page_number: int,
    excerpt: str,
    heuristic_confidence: float,
    block_number: int,
    ocr_confidence: float | None = None,
    *,
    line_number: int | None = None,
    bbox: tuple[float, float, float, float] | None = None,
    source: str | None = None,
    field_status: str = "found",
    decision_comment: str = "",
) -> None:
    # La confianza final es la más baja entre la de la regla y la del OCR.
    confidence = min(heuristic_confidence, ocr_confidence) if ocr_confidence is not None else heuristic_confidence
    item = FieldEvidence(
        observed_value=observed.strip(),
        normalized_value=normalized,
        page_number=page_number,
        excerpt=excerpt,
        confidence=round(confidence, 4),
        block_number=block_number,
        line_number=line_number,
        bbox=bbox,
        source=source or result.reading_source,
        ocr_confidence=round(ocr_confidence, 4) if ocr_confidence is not None else None,
        heuristic_confidence=round(heuristic_confidence, 4),
        field_status=field_status,
        decision_comment=decision_comment,
    )
    result.field_evidence.setdefault(field_name, []).append(item)
    result.evidence.setdefault(field_name, []).append(excerpt)
    result.fields_original.setdefault(field_name, observed.strip())
    result.field_statuses[field_name] = field_status
    result.confidence[field_name] = min(result.confidence.get(field_name, 1.0), round(confidence, 4))


# Reconoce documentos de soporte que NO son comprobantes de pago (consulta de validez, asiento SAP, guía de remisión,
# estado de cuenta, HES, orden de compra, anexo de factura negociable) por frases propias de cada uno.
# Devuelve (tipo, regla), o (None, None) si no reconoce ninguno.
def _detect_non_receipt_type(lines: list[str]) -> tuple[str | None, str | None]:
    """Identifica documentos de soporte no fiscales a partir de señales locales del bloque."""
    joined = "\n".join(lines)
    upper = joined.upper()

    # 1. Constancia de consulta de validez de CPE en SUNAT
    if (
        "CONSULTA DE VALIDEZ DEL CPE" in upper
        or "CONSULTA VALIDEZ DEL COMPROBANTE DE PAGO" in upper
        or ("RESULTADO DE LA CONSULTA" in upper and "ES UN COMPROBANTE DE PAGO V" in upper)
    ):
        return "constancia_consulta_validez_cpe_sunat", "non_receipt_sunat_cpe_query"

    # 2. Reporte / asiento contable SAP
    if (
        "VISUALIZAR DOCUMENTO : VISTA DE ENTRADA" in upper
        or "CL.DOC. : RE ( FACTURA BRUTO )" in upper
        or ("SOCIEDAD" in upper and "FE.CONTAB." in upper and "CLAVE CONTAB" in upper)
    ):
        return "reporte_asiento_contable_sap", "non_receipt_sap_accounting_doc"

    # 3. Guía de remisión (remitente o transportista)
    has_receipt_header = bool(
        re.search(r"\b(?:FACTURA\s+ELECTR[ÓO]NICA|BOLETA\s+DE\s+VENTA|NOTA\s+DE\s+CR[EÉ]DITO|NOTA\s+DE\s+D[EÉ]BITO)\b", upper)
    )
    if not has_receipt_header and re.search(
        r"\bGU[ÍI]A\s+DE\s+REMISI[ÓO]N(?:\s+ELECTR[ÓO]NICA|\s+REMITENTE|\s+TRANSPORTISTA)?\b",
        upper,
    ):
        return "guia_de_remision", "non_receipt_remision_guide"

    # 4. Estado de cuenta bancario
    if not has_receipt_header and (
        "ACCOUNT STATEMENT INQUIRY" in upper
        or "CONSULTA DE DECLARACIÓN DE CUENTA" in upper
        or "CONSULTA DE DECLARACION DE CUENTA" in upper
        or "ESTADO DE CUENTA CORRIENTE" in upper
        or "ESTADO DE CUENTA BANCARIO" in upper
        or "ESTADO DE CUENTA" in upper
        or ("OPENING BALANCE" in upper and "CLOSING BALANCE" in upper)
    ):
        return "estado_de_cuenta_bancario", "non_receipt_bank_statement"

    # 5. Hoja de Entrada de Servicios (HES) o flujo de liberación/aprobación HES en SAP
    if not has_receipt_header and (
        "HOJA DE ENTRADA DE SERVICIOS" in upper
        or ("HOJA DE ENTRADA" in upper and ("SERVICIOS/FECHA" in upper or "SERV ARIBA" in upper))
        or "ACEPTACIÓN DE SERVICIOS" in upper
        or "ACEPTACION DE SERVICIOS" in upper
        or "N° HOJA ENTRADA SERV" in upper
        or "H.ENTR.SERV." in upper
        or ("RUTA DE APROBACIÓN" in upper and ("LIBERARSE" in upper or "LIBERADO EN" in upper))
        or ("RUTA DE APROBACION" in upper and ("LIBERARSE" in upper or "LIBERADO EN" in upper))
    ):
        return "hoja_entrada_servicios_hes", "non_receipt_service_entry_sheet"

    # 6. Orden de compra / Pedido
    if not has_receipt_header and (
        re.search(r"\b(?:ORDEN\s+DE\s+COMPRA|PURCHASE\s+ORDER|ANEXO\s+DE\s+PEDIDO)\b", upper)
        or (re.search(r"\bPEDIDO\s*:\s*[0-9]{6,15}\b", upper) and ("PROVEEDOR" in upper or "COMPRADOR" in upper or "PÁGINA:" in upper or "PAGINA:" in upper))
        or ("GENERAL CONTRACTING TERMS AND CONDITIONS APPLYING TO PURCHASE ORDERS" in upper)
    ):
        return "orden_de_compra_pedido", "non_receipt_purchase_order"

    # 7. Anexo de factura negociable (página de endoso sin numeración propia)
    if (
        "INFORMACIÓN ADICIONAL DE LA FACTURA NEGOCIABLE" in upper
        or "INFORMACION ADICIONAL DE LA FACTURA NEGOCIABLE" in upper
        or "FECHA DE ENTREGA DE LA FACTURA NEGOCIABLE" in upper
    ) and not re.search(r"\b[FBE][A-Z0-9]{2,3}\s*(?:N(?:RO|°|º|\.)\s*)?[-–]\s*[0-9]{1,20}\b", upper):
        return "anexo_factura_negociable", "continuation_negotiable_invoice"

    return None, None


# Busca en las líneas el título de algún tipo de comprobante. Acepta el título solo o con un prefijo de membrete
# (RUC, teléfono, nombre del emisor) y un sufijo («ELECTRÓNICA», serie y número).
# Devuelve (línea, código, nombre, texto de la línea).
def _type_hits(lines: list[str], catalog: Mapping[str, str]) -> list[tuple[int, str, str, str]]:
    hits: list[tuple[int, str, str, str]] = []
    for index, line in enumerate(lines):
        if _NON_HEADER_TYPE_CONTEXT.search(line) or re.match(
            r"^\s*(?:GLOSA|SUSTENTO|OBSERVACIONES|SON|DETALLE|DESCRIPCI[ÓO]N)\b", line, re.IGNORECASE
        ):
            continue
        for code, label, pattern in _TYPE_PATTERNS:
            if code not in catalog or not (match := pattern.search(line)):
                continue
            prefix = line[:match.start()].strip(" \t:|–-")
            suffix = line[match.end():]
            if (not prefix or _HEADER_CONTEXT.search(prefix)) and _HEADER_SUFFIX.fullmatch(suffix):
                hits.append((index, code, label, line.strip()))
    return hits


# ¿Este texto contiene el título de ese tipo de comprobante?
def has_document_title(text: str, document_type: str) -> bool:
    return any(code == document_type for _, code, _, _ in _type_hits(text.splitlines(), DEFAULT_DOCUMENT_TYPES))


def document_title_from_line(line: str, document_type: str | None) -> str:
    """Devuelve la denominación observada dentro de la línea, sin el texto vecino.

    En muchos comprobantes el título comparte renglón con el teléfono, el RUC o el nombre
    del emisor porque están en columnas distintas de la página. Aquí nos quedamos con el
    tramo que corresponde al título; la línea completa se conserva aparte como evidencia.
    No se completa ni se corrige ningún carácter: se recorta, nada más.
    """
    text = (line or "").strip()
    pattern = next((p for code, _, p in _TYPE_PATTERNS if code == document_type), None)
    if pattern is None or not (match := pattern.search(text)):
        return text
    title = match.group().strip()
    if tail := _TITLE_TAIL.match(text[match.end():]):
        title = f"{title} {tail.group(1).strip()}"
    return title


# Encuentra los números que parecen un RUC: 11 caracteres, con al menos 8 dígitos y cerca de un rótulo «RUC».
def _ruc_mentions(lines: list[str]) -> list[tuple[int, str]]:
    mentions: list[tuple[int, str]] = []
    for index, line in enumerate(lines):
        for match in _RUC_TOKEN.finditer(line):
            token = match.group(1)
            if sum(character.isdigit() for character in token) < 8:
                continue
            # Un número de 11 dígitos sin un rótulo «RUC» cerca puede ser una cuenta, un contrato o una autorización: se descarta.
            nearby = lines[max(0, index - 1) : min(len(lines), index + 2)]
            if any(_RUC_LABEL.search(item) for item in nearby):
                # Si hay varios números en la misma línea (por ejemplo, un RUC y un teléfono), se queda con el que empieza por
                # 10, 15, 17 o 20, o con el más cercano al rótulo.
                if token[:2] in {"10", "15", "17", "20"} or _RUC_LABEL.search(line[: match.start() + 15]):
                    mentions.append((index, token))
    return mentions


# ¿Esta serie tiene forma de serie fiscal? Letra F, B o E con 3 o 4 caracteres y algún dígito;
# o 3 o 4 dígitos, si el contexto es el de un comprobante.
def _is_valid_series_token(series: str, contextual: bool) -> bool:
    up = series.upper()
    if up.isdigit():
        return len(up) in {3, 4} and contextual
    if len(up) in {3, 4} and up[0] in {"F", "B", "E"} and any(ch.isdigit() for ch in up):
        return True
    return False


# Busca series y correlativos en las líneas. Devuelve dos listas: las identidades propias del comprobante
# y los comprobantes a los que hace referencia (los que modifica una nota, por ejemplo).
def _serial_signals(lines: list[str]) -> tuple[list[tuple[str, str, int, tuple[int, ...]]], list[tuple[str, int]]]:
    """Devuelve (identidades propias del comprobante, series-número referenciados)."""
    found: list[tuple[str, str, int, tuple[int, ...]]] = []
    referenced: list[tuple[str, int]] = []

    for index, line in enumerate(lines):
        # Paso 1: serie y número juntos («F001-00003080»). Se descartan los que forman parte de una fecha,
        # de un teléfono o de un comprobante referenciado.
        date_spans = [match.span() for match in _DATE.finditer(line)]
        is_ref_line = bool(_REFERENCE_SERIAL_LINE.search(line))
        if index > 0 and re.search(r"\b(?:DOCUMENTO\s+DE\s+REFERENCIA|DOCUMENTO\s+QUE\s+MODIFICA)\b", lines[index - 1], re.IGNORECASE):
            is_ref_line = True
        if index > 1 and re.search(r"\b(?:DOCUMENTO\s+DE\s+REFERENCIA|DOCUMENTO\s+QUE\s+MODIFICA)\b", " ".join(lines[max(0, index - 3) : index]), re.IGNORECASE):
            if "FACTURA" in " ".join(lines[max(0, index - 2) : index + 1]).upper() or "TIPO DE NOTA" in line.upper():
                is_ref_line = True

        contextual = bool(
            re.search(
                r"\b(?:FACTURA|BOLETA|NOTA|SERIE|NRO\.?|N[ÚU]MERO|CORRELATIVO|COMPROBANTE)\b",
                line,
                re.IGNORECASE,
            )
        )
        for match in _SERIAL_PAIR.finditer(line):
            if any(match.start() < end and match.end() > start for start, end in date_spans):
                continue
            prefix = line[: match.start()]
            if re.search(r"\b(?:TEL(?:[ÉE]F(?:ONO)?|\.)?|FAX|ANEXO|CEL(?:ULAR)?)\s*[:.#-]?\s*(?:\([0-9]+\)\s*)?$", prefix, re.IGNORECASE):
                continue
            series = match.group("series").upper()
            number = match.group("number")
            if not _is_valid_series_token(series, contextual):
                continue
            if is_ref_line:
                referenced.append((f"{series}-{number}", index))
                continue
            found.append((series, number, index, (index,)))

    # Paso 2: serie y número escritos aparte, cada uno con su rótulo («Serie: …», «Nro: …»).
    # Sólo se emparejan si hay uno de cada y están a menos de tres líneas.
    series_labels: list[tuple[str, int]] = []
    number_labels: list[tuple[str, int]] = []
    for index, line in enumerate(lines):
        if _REFERENCE_SERIAL_LINE.search(line):
            continue
        series_labels.extend((match.group(1), index) for match in _SERIES_LABEL.finditer(line))
        # Se evitan los «N°» que no son el número del comprobante: RUC, ley, decreto, resolución, direcciones o contratos.
        if re.search(r"\b(?:R\.?U\.?C\.?|LEY|DECRETO|RESOLUCI[ÓO]N|ART[ÍI]CULO|URB\.|AV\.|CARR\.|DIRECCI[ÓO]N|RAUC|CONTRATO|PROVEEDOR)\b", line, re.IGNORECASE):
            if not re.search(r"\bCORRELATIVO\b", line, re.IGNORECASE):
                continue
        for match in _NUMBER_LABEL.finditer(line):
            # Se salta si ese número ya forma parte de una serie-número de la misma línea.
            if any(idx == index and num == match.group(1) for _, num, idx, _ in found):
                continue
            number_labels.append((match.group(1), index))

    if len(series_labels) == 1 and len(number_labels) == 1 and abs(series_labels[0][1] - number_labels[0][1]) <= 3:
        series, series_line = series_labels[0]
        number, number_line = number_labels[0]
        found.append((series, number, series_line, tuple(sorted({series_line, number_line}))))
    elif not found and (series_labels or number_labels):
        for value, index in series_labels:
            found.append((value, "", index, (index,)))
        for value, index in number_labels:
            found.append(("", value, index, (index,)))

    # Paso 3: quitar los repetidos.
    unique: dict[tuple[str, str], tuple[int, tuple[int, ...]]] = {}
    for series, number, index, evidence_lines in found:
        key = (series.upper(), number)
        unique.setdefault(key, (index, evidence_lines))
    return (
        [(series, number, index, evidence) for (series, number), (index, evidence) in unique.items()],
        referenced,
    )


# La «huella» de un bloque: los tipos de comprobante y las series-número que contiene.
# Sirve para decidir si un trozo de página puede considerarse un comprobante por sí solo.
def _block_signature(lines: list[str], catalog: Mapping[str, str]) -> tuple[tuple[str, ...], tuple[tuple[str, str], ...]]:
    types = tuple(sorted({code for _, code, _, _ in _type_hits(lines, catalog)}))
    signals, _ = _serial_signals(lines)
    serials = tuple(sorted({(series.upper(), number) for series, number, _, _ in signals if series and number}))
    return types, serials


# Términos de soporte que también se usan como rótulo en la cabecera de muchas facturas.
_FIELD_LABEL_SUPPORT_TERM = re.compile(
    r"\b(?:ORDEN\s+DE\s+COMPRA|PURCHASE\s+ORDER|GU[ÍI]A\s+DE\s+REMISI[ÓO]N)\b",
    re.IGNORECASE,
)


def _support_block_stands_alone(segment: Sequence[str], title_index: int) -> bool:
    """¿El bloque de soporte se sostiene sin su línea de título?

    «Orden de compra» y «Guía de remisión» son, además de documentos, rótulos de la cabecera
    de muchas facturas. Cuando el valor del rótulo viene vacío, el OCR puede perder los dos
    puntos y dejarlo solo en su renglón, con aspecto de encabezado de otro documento. En ese
    caso pedimos que el resto del bloque se clasifique como soporte por sí mismo; si no lo
    hace, lo que hay debajo pertenece al comprobante y cortar por ahí dejaría fuera sus
    importes. Los títulos que no se usan como rótulo (un estado de cuenta, una HES) siguen
    cortando como antes.
    """
    title = segment[title_index] if 0 <= title_index < len(segment) else ""
    match = _FIELD_LABEL_SUPPORT_TERM.search(title)
    if not match or title[match.end():].strip(" \t:|–-"):
        return True
    body = [line for index, line in enumerate(segment) if index != title_index and line.strip()]
    return bool(body) and _detect_non_receipt_type(body)[0] is not None


def _split_blocks(lines: list[str], catalog: Mapping[str, str]) -> list[list[str]]:
    """Segmenta una página antes de clasificar cuando coexisten varios CDP o un bloque de soporte y un CDP."""
    # Paso 1: dónde están los títulos de comprobante y los encabezados de soporte.
    hits = _type_hits(lines, catalog)
    header_indexes = sorted({index for index, _, _, _ in hits})
    # Una etiqueta dentro de una factura (por ejemplo, «Guía de remisión:»)
    # no abre otro documento ni corta el bloque de importes.
    support_indexes = [
        idx for idx, line in enumerate(lines)
        if (match := _SUPPORT_HEADER_LINE_RE.search(line))
        and not line[:match.start()].strip(" \t|–-")
        and not line[match.end():].lstrip().startswith(":")
    ]

    # Caso 1: un encabezado de soporte (p. ej. estado bancario o HES) antecede o sigue a un comprobante fiscal
    if len(header_indexes) == 1 and support_indexes:
        header_idx = header_indexes[0]
        rec_start = header_idx
        for candidate in range(max(0, header_idx - 3), header_idx):
            if _SUPPORT_HEADER_LINE_RE.search(lines[candidate]):
                continue
            if _RUC_LABEL.search(lines[candidate]) or _HEADER_COMPANY_CUE.search(lines[candidate]):
                rec_start = candidate
                break
        before_support = [s for s in support_indexes if s < rec_start]
        if before_support and rec_start > 0:
            rec_segment = lines[rec_start:]
            rec_types, rec_serials = _block_signature(rec_segment, catalog)
            if (
                rec_types
                and (rec_serials or _ruc_mentions(rec_segment))
                and _support_block_stands_alone(lines[:rec_start], before_support[0])
            ):
                return [lines[:rec_start], rec_segment]
        after_support = [s for s in support_indexes if s > header_idx + 2]
        if after_support:
            split_at = after_support[0]
            rec_segment = lines[:split_at]
            rec_types, rec_serials = _block_signature(rec_segment, catalog)
            if (
                rec_types
                and (rec_serials or _ruc_mentions(rec_segment))
                and _support_block_stands_alone(lines[split_at:], 0)
            ):
                return [rec_segment, lines[split_at:]]

    # Caso 2: con menos de dos títulos no hay nada que partir; la página es un solo bloque.
    if len(header_indexes) < 2:
        return [lines]

    # Caso 3: varios títulos en la misma página. Cada uno abre un bloque, retrocediendo hasta tres líneas
    # para incluir su membrete (RUC o nombre del emisor).
    starts = [0]
    for header_index in header_indexes[1:]:
        start = header_index
        for candidate in range(max(starts[-1], header_index - 3), header_index):
            if _RUC_LABEL.search(lines[candidate]) or (
                _RUC_TOKEN.search(lines[candidate]) and candidate > 0 and _RUC_LABEL.search(lines[candidate - 1])
            ):
                start = candidate
        if start > starts[-1]:
            starts.append(start)
    if len(starts) < 2:
        return [lines]

    segments = [
        lines[start : (starts[index + 1] if index + 1 < len(starts) else len(lines))]
        for index, start in enumerate(starts)
    ]
    # Sólo se parte si cada trozo tiene su propio tipo y su propia serie-número, y las identidades son distintas;
    # si no, la página se queda entera.
    signatures = [_block_signature(segment, catalog) for segment in segments]
    if any(not types or not serials for types, serials in signatures):
        return [lines]
    all_serials = [serial for _, serials in signatures for serial in serials]
    if len(set(all_serials)) < 2:
        return [lines]
    return segments


# Dónde empieza el bloque del cliente (receptor): la primera línea con un rótulo de cliente o «Razón social:».
def _client_section_start(lines: list[str]) -> int | None:
    """Devuelve el índice de la primera línea que abre el bloque del cliente o receptor."""
    for idx, line in enumerate(lines):
        if _RECIPIENT_ROLE.search(line):
            return idx
        if re.search(r"^\s*RAZ[ÓO]N\s+SOCIAL\s*:", line, re.IGNORECASE) and idx >= 4:
            return idx
    return None


# Decide el rol de un RUC (emisor, receptor o tercero) por dónde está y por las etiquetas que lo rodean.
# Si no puede decidirlo con seguridad, devuelve None y el RUC queda como candidato.
def _role_for_ruc(
    lines: list[str],
    index: int,
    token: str,
    target_recipient_ruc: str | None,
) -> tuple[str | None, str]:
    """Determina si un RUC corresponde al emisor, al adquirente o a un tercero según su posición y etiquetas."""
    line = lines[index]
    prev_line = lines[index - 1] if index > 0 else ""

    # 1. Rol explícito de tercero (transportista, banco, garante, factoring) en la misma línea o la inmediata anterior
    if _THIRD_PARTY_ROLE.search(line) or (_THIRD_PARTY_ROLE.search(prev_line) and not _ISSUER_ROLE.search(line) and not _RECIPIENT_ROLE.search(line)):
        return "third_party_ruc", "RUC asociado a rol explícito de tercero (transporte/banco/cesionario)"

    if re.match(r"^\s*NOMBRE\s*/\s*RAZ[ÓO]N\s+SOCIAL\s*:", line, re.IGNORECASE) and _RUC_LABEL.search(line):
        return "recipient_ruc", "RUC en la fila de nombre o razón social del comprador"

    # 2. Etiqueta explícita de emisor/proveedor en la misma línea o en las líneas previas inmediatas
    issuer_preceding = [
        index - other
        for other in range(max(0, index - 4), index + 1)
        if _ISSUER_ROLE.search(lines[other])
    ]
    # Una etiqueta de cliente/adquirente introduce los datos situados en esa misma línea o debajo de ella,
    # nunca un RUC que apareció líneas más arriba en la cabecera del emisor.
    recipient_preceding = [
        index - other
        for other in range(max(0, index - 2), index + 1)
        if _RECIPIENT_ROLE.search(lines[other])
    ]
    issuer_distance = min(issuer_preceding, default=99)
    recipient_distance = min(recipient_preceding, default=99)

    if min(issuer_distance, recipient_distance) <= 4 and issuer_distance != recipient_distance:
        if issuer_distance < recipient_distance:
            return "emitter_ruc", "Etiqueta explícita de emisor/proveedor junto al RUC"
        if recipient_distance <= 2:
            return "recipient_ruc", "Etiqueta explícita de cliente/adquirente/señor(es) en la línea o líneas previas"

    # 3. Por su posición respecto del bloque del cliente: dentro de él es el receptor; en la cabecera, antes de él, el emisor.
    client_start = _client_section_start(lines)
    if client_start is not None:
        if client_start <= index <= client_start + 5:
            return "recipient_ruc", "RUC ubicado dentro del bloque de datos del cliente/receptor"
        header_slice = lines[:client_start]
        client_slice = lines[client_start : min(len(lines), client_start + 6)]
        client_has_own_ruc = any(_RUC_TOKEN.search(cl) for cl in client_slice)
        has_header_structure = any(_HEADER_COMPANY_CUE.search(l) for l in header_slice) or (
            client_has_own_ruc and any(_SERIAL_PAIR.search(l) or _SERIES_LABEL.search(l) for l in header_slice)
        )
        if index < client_start and index <= 14:
            if has_header_structure:
                if target_recipient_ruc and token == target_recipient_ruc and not client_has_own_ruc:
                    return None, "RUC coincide con receptor objetivo pero aparece en cabecera sin RUC separado en sección cliente"
                return "emitter_ruc", "RUC ubicado en cabecera del emisor antes de la sección explícita del cliente"
        if index > client_start + 5:
            # Recuadro fiscal SUNAT (p. ej. FACTURA ELECTRONICA / RUC: <11 digitos> / E001-123),
            # que en diseños a dos columnas puede quedar leído por OCR después del bloque del cliente.
            box_window = lines[max(0, index - 2) : min(len(lines), index + 3)]
            has_box_type = any(
                any(p.search(w_line) for _, _, p in _TYPE_PATTERNS) and not _NON_HEADER_TYPE_CONTEXT.search(w_line)
                for w_line in box_window
            )
            has_box_serial = any(
                _SERIAL_PAIR.search(w_line) and not _REFERENCE_SERIAL_LINE.search(w_line)
                for w_line in box_window
            )
            if has_box_type and has_box_serial and client_has_own_ruc and has_header_structure:
                return "emitter_ruc", "RUC ubicado dentro del recuadro fiscal junto al tipo de comprobante y serie-correlativo"
            return "third_party_ruc", "RUC observado fuera de la cabecera del emisor y fuera del bloque del adquirente"

    # Si no se detectó sección de cliente pero la cabecera tiene datos societarios/domicilio del emisor
    if index <= 8 and any(_HEADER_COMPANY_CUE.search(l) for l in lines[: min(len(lines), 10)]):
        if not (target_recipient_ruc and token == target_recipient_ruc):
            return "emitter_ruc", "RUC ubicado en membrete/cabecera fiscal con datos societarios/domicilio del emisor"

    return None, "RUC sin rol inequívoco en el bloque; se conserva como candidato"


# Para páginas de dos columnas: empareja el rótulo «RUC» con su valor cuando ambos están en la misma fila
# del bloque del comprador, usando sus coordenadas.
def _recipient_rucs_in_columns(layout: Sequence[LayoutLine]) -> dict[tuple[int, str], tuple[int, int]]:
    """Une una etiqueta RUC a su valor en la misma fila del bloque del comprador."""
    result = {}
    for label_index, label in enumerate(layout):
        if not label.bbox or not _RUC_LABEL.fullmatch(label.text.strip(" :.")):
            continue
        left, top, right, bottom = label.bbox
        height = bottom - top
        if height <= 0:
            continue
        anchors = [
            index for index, anchor in enumerate(layout)
            if anchor.source == label.source and anchor.bbox
            and re.match(r"^\s*(?:SE[ÑN]OR|SEFIOR|CLIENTE|ADQUIRENTE|RECEPTOR|COMPRADOR)\b", anchor.text, re.IGNORECASE)
            and 0 < top - anchor.bbox[1] <= 4 * height
            and abs(left - anchor.bbox[0]) <= 2 * height
        ]
        if not anchors:
            continue
        matches = []
        for index, value in enumerate(layout):
            if value.source != label.source or not value.bbox:
                continue
            token = value.text.strip(" .:|")
            if not re.fullmatch(r"[A-Z0-9]{11}", token, re.IGNORECASE):
                continue
            if value.bbox[0] >= right and abs((top + bottom) - (value.bbox[1] + value.bbox[3])) <= height:
                matches.append((index, token))
        if len(matches) == 1:
            result[matches[0]] = (anchors[-1], label_index)
    return result


# Encuentra los RUC del bloque, valida cada uno (formato y dígito de control) y le asigna un rol.
# Si por rol hay un único valor válido, lo da por leído; si hay varios o alguno dudoso, los deja como candidatos para revisión.
def _extract_rucs(
    result: ExtractionResult,
    lines: list[str],
    page_number: int,
    block_number: int,
    line_confidences: Mapping[str, float],
    layout_lines: Sequence[LayoutLine] | None,
    target_recipient_ruc: str | None,
) -> bool:
    # Paso 1: preparar los contenedores por rol.
    role_values: dict[str, list[tuple[str, int, bool, str]]] = {
        "emitter_ruc": [],
        "recipient_ruc": [],
        "third_party_ruc": [],
    }
    unassigned: list[tuple[str, int]] = []
    aligned_layout = (
        layout_lines if layout_lines and len(layout_lines) == len(lines)
        and all(item.text.strip() == line for item, line in zip(layout_lines, lines)) else []
    )
    column_matches = _recipient_rucs_in_columns(aligned_layout)
    # Paso 2: reunir los RUC, tanto los que están junto a un rótulo como los alineados en columnas.
    mentions = sorted(set(_ruc_mentions(lines)) | column_matches.keys())

    # Paso 3: para cada RUC, decidir su rol, validarlo y guardar su evidencia.
    for index, token in mentions:
        if (index, token) in column_matches:
            role, comment = "recipient_ruc", "RUC alineado con su etiqueta en el bloque del comprador"
        else:
            role, comment = _role_for_ruc(lines, index, token, target_recipient_ruc)
        ruc_val = validate_ruc(token)
        valid = bool(token.isdigit() and ruc_val.format_valid and ruc_val.checksum_valid)
        nearby_indexes = list(range(max(0, index - 2), min(len(lines), index + 3)))
        excerpt = _excerpt(lines, [index, *nearby_indexes], radius=0)
        if (index, token) in column_matches:
            excerpt = "\n".join(lines[position] for position in (*column_matches[index, token], index))
        raw_confidence = _line_confidence(lines, [index], line_confidences, layout_lines)
        bbox = _line_bbox([index], layout_lines)
        line_no = layout_lines[index].line_number if layout_lines and index < len(layout_lines) else index + 1

        if role == "third_party_ruc":
            _add_field(
                result,
                "third_party_ruc",
                token,
                token if valid else None,
                page_number,
                excerpt,
                0.90 if valid else 0.55,
                block_number,
                raw_confidence,
                line_number=line_no,
                bbox=bbox,
                field_status="found" if valid else "invalid_checksum_or_format",
                decision_comment=comment,
            )
            role_values["third_party_ruc"].append((token, index, valid, comment))
            if token not in result.third_party_rucs:
                result.third_party_rucs.append(token)
        elif role in {"emitter_ruc", "recipient_ruc"}:
            _add_field(
                result,
                role,
                token,
                token if valid else None,
                page_number,
                excerpt,
                0.96 if valid else 0.55,
                block_number,
                raw_confidence,
                line_number=line_no,
                bbox=bbox,
                field_status="found" if valid else "invalid_checksum_or_format",
                decision_comment=comment,
            )
            role_values[role].append((token, index, valid, comment))
        else:
            unassigned.append((token, index))
            _add_field(
                result,
                "emitter_ruc",
                token,
                None,
                page_number,
                excerpt,
                0.55,
                block_number,
                raw_confidence,
                line_number=line_no,
                bbox=bbox,
                field_status="ambiguous_role",
                decision_comment=comment,
            )
            if valid:
                result.candidates.setdefault("emitter_ruc", []).append(token)
            if target_recipient_ruc and token == target_recipient_ruc:
                result.candidates.setdefault("recipient_ruc", []).append(token)

    # Paso 4: por rol, un único valor válido se acepta; si hay varios o alguno inválido, quedan como candidatos.
    for field_name in ("emitter_ruc", "recipient_ruc"):
        values = role_values[field_name]
        distinct = {value for value, _, valid, _ in values if valid}
        invalid = {value for value, _, valid, _ in values if not valid}
        if len(distinct) == 1 and not invalid:
            setattr(result, field_name, next(iter(distinct)))
        elif distinct or invalid:
            # Conservamos el literal observado (sin sustituir S->5 ni O->0) en candidates para revisión.
            result.candidates.setdefault(field_name, []).extend(sorted(distinct | invalid))
            result.observations.append(f"RUC {field_name} observado con valor dudoso/inválido o más de un candidato")

    result.candidates["emitter_ruc"] = sorted(set(result.candidates.get("emitter_ruc", [])))
    result.candidates["recipient_ruc"] = sorted(set(result.candidates.get("recipient_ruc", [])))
    if result.third_party_rucs:
        result.candidates["third_party_ruc"] = sorted(set(result.third_party_rucs))

    if result.emitter_ruc:
        result.issuer_role_status = "verified_issuer"
    elif result.candidates.get("emitter_ruc"):
        result.issuer_role_status = "candidate_unassigned"
    else:
        result.issuer_role_status = "not_found"

    if result.recipient_ruc:
        if target_recipient_ruc and result.recipient_ruc == target_recipient_ruc:
            result.recipient_ruc_match = True
            result.recipient_role_status = "verified_recipient"
        elif target_recipient_ruc:
            result.recipient_ruc_match = False
            result.recipient_role_status = "mismatch"
        else:
            result.recipient_ruc_match = True
            result.recipient_role_status = "verified_recipient"
    elif target_recipient_ruc and (
        target_recipient_ruc in result.candidates.get("recipient_ruc", [])
        or target_recipient_ruc in result.candidates.get("emitter_ruc", [])
    ):
        result.recipient_ruc_match = False
        result.recipient_role_status = "candidate_unassigned"
    else:
        result.recipient_ruc_match = False
        result.recipient_role_status = "not_found"

    if unassigned:
        result.observations.append("RUC sin rol inequívoco en el bloque; se conserva como candidato")
    return bool(mentions)


# Detecta el título del comprobante (factura, boleta, nota...). Con un único tipo lo da por leído, guardando el título
# recortado como valor observado y la línea completa como evidencia; con varios incompatibles, los deja como candidatos.
def _extract_types(
    result: ExtractionResult,
    lines: list[str],
    catalog: Mapping[str, str],
    page_number: int,
    block_number: int,
    line_confidences: Mapping[str, float],
    layout_lines: Sequence[LayoutLine] | None,
) -> set[str]:
    hits = _type_hits(lines, catalog)
    codes = {code for _, code, _, _ in hits}
    note_codes = codes & {"07", "08"}
    selected = note_codes if len(note_codes) == 1 else codes
    if len(selected) == 1:
        code = next(iter(selected))
        relevant = [hit for hit in hits if hit[1] == code]
        result.document_type = code
        # Guardamos la denominación recortada como valor observado y la línea entera como evidencia.
        header_line = relevant[0][3]
        observed = document_title_from_line(header_line, code)
        idx = relevant[0][0]
        _add_field(
            result,
            "document_type",
            observed,
            code,
            page_number,
            header_line,
            0.94,
            block_number,
            _line_confidence(lines, [hit[0] for hit in relevant], line_confidences, layout_lines),
            line_number=layout_lines[idx].line_number if layout_lines and idx < len(layout_lines) else idx + 1,
            bbox=_line_bbox([hit[0] for hit in relevant], layout_lines),
            decision_comment=f"Tipo documental {code} ({catalog.get(code, '')}) detectado en encabezado",
        )
    elif selected:
        result.candidates["document_type"] = sorted(selected)
        result.observations.append("se encontraron tipos documentales incompatibles en el mismo bloque")
        for index, code, _, line in hits:
            if code in selected:
                _add_field(
                    result,
                    "document_type",
                    document_title_from_line(line, code),
                    code,
                    page_number,
                    line,
                    0.55,
                    block_number,
                    _line_confidence(lines, [index], line_confidences, layout_lines),
                    line_number=layout_lines[index].line_number if layout_lines and index < len(layout_lines) else index + 1,
                    bbox=_line_bbox([index], layout_lines),
                    field_status="conflicting_candidates",
                )
    return selected


# Detecta serie y correlativo. Un único par completo se acepta; varios pares distintos quedan como candidatos;
# si sólo hay mitades, se guardan por separado. También anota el comprobante referenciado en las notas.
def _extract_serials(
    result: ExtractionResult,
    lines: list[str],
    page_number: int,
    block_number: int,
    line_confidences: Mapping[str, float],
    layout_lines: Sequence[LayoutLine] | None,
) -> list[tuple[str, str]]:
    signals, referenced = _serial_signals(lines)
    if referenced and not result.referenced_document:
        ref_val, ref_idx = referenced[0]
        result.referenced_document = ref_val
        _add_field(
            result,
            "referenced_document",
            ref_val,
            ref_val,
            page_number,
            lines[ref_idx].strip(),
            0.90,
            block_number,
            _line_confidence(lines, [ref_idx], line_confidences, layout_lines),
            line_number=layout_lines[ref_idx].line_number if layout_lines and ref_idx < len(layout_lines) else ref_idx + 1,
            bbox=_line_bbox([ref_idx], layout_lines),
            decision_comment="Comprobante de referencia / modificado citado en el documento",
        )

    complete = [(series, number, index, indexes) for series, number, index, indexes in signals if series and number]
    unique_pairs = list(dict.fromkeys((series.upper(), number) for series, number, _, _ in complete))
    if len(unique_pairs) == 1:
        series, number = unique_pairs[0]
        matching = next(item for item in complete if (item[0].upper(), item[1]) == (series, number))
        _, _, primary_idx, indexes = matching
        excerpt = _excerpt(lines, indexes)
        score = _line_confidence(lines, indexes, line_confidences, layout_lines)
        bbox = _line_bbox(indexes, layout_lines)
        line_no = layout_lines[primary_idx].line_number if layout_lines and primary_idx < len(layout_lines) else primary_idx + 1
        normalized = normalize_correlativo(number)
        result.series = series
        result.correlativo_original = number
        result.correlativo_normalized = normalized.normalized
        _add_field(
            result,
            "series",
            matching[0],
            series,
            page_number,
            excerpt,
            0.92,
            block_number,
            score,
            line_number=line_no,
            bbox=bbox,
        )
        _add_field(
            result,
            "correlativo",
            number,
            normalized.normalized,
            page_number,
            excerpt,
            0.94 if normalized.valid else 0.55,
            block_number,
            score,
            line_number=line_no,
            bbox=bbox,
            field_status="found" if normalized.valid else "invalid_length_or_format",
        )
        if not normalized.valid:
            result.observations.append(normalized.reason or "correlativo no normalizable")
        # Una letra en esta posición puede ser válida en texto nativo. Si proviene
        # de OCR, conservamos la lectura y pedimos revisar el carácter en la imagen.
        if result.reading_source == "ocr" and re.fullmatch(r"[FBE](?=[0-9OI]*[0-9])(?=[0-9OI]*[OI])[0-9OI]{2,3}", series):
            if "series" not in result.secondary_review_fields:
                result.secondary_review_fields.append("series")
            result.observations.append(
                f"posible confusión O/0 o I/1 en serie ({series}); se conserva el literal para revisión"
            )
    elif len(unique_pairs) > 1:
        result.candidates["document_identity"] = [f"{series}-{number}" for series, number in unique_pairs]
        result.candidates["series"] = sorted({series for series, _ in unique_pairs})
        result.candidates["correlativo"] = sorted({number for _, number in unique_pairs})
        result.observations.append("hay varias identidades en el mismo bloque; no se emparejan globalmente")
        for series, number, index, indexes in complete:
            excerpt = _excerpt(lines, indexes)
            score = _line_confidence(lines, indexes, line_confidences, layout_lines)
            bbox = _line_bbox(indexes, layout_lines)
            line_no = layout_lines[index].line_number if layout_lines and index < len(layout_lines) else index + 1
            _add_field(
                result,
                "series",
                series,
                series.upper(),
                page_number,
                excerpt,
                0.55,
                block_number,
                score,
                line_number=line_no,
                bbox=bbox,
                field_status="conflicting_candidates",
            )
            _add_field(
                result,
                "correlativo",
                number,
                normalize_correlativo(number).normalized,
                page_number,
                excerpt,
                0.55,
                block_number,
                score,
                line_number=line_no,
                bbox=bbox,
                field_status="conflicting_candidates",
            )
    else:
        partial_series = [series for series, number, _, _ in signals if series and not number]
        partial_numbers = [number for series, number, _, _ in signals if number and not series]
        if len(set(partial_series)) == 1:
            observed = partial_series[0]
            result.series = observed.upper()
            result.candidates["correlativo"] = sorted(set(partial_numbers))
            idx = next(index for series, number, index, _ in signals if series == observed and not number)
            _add_field(
                result,
                "series",
                observed,
                observed.upper(),
                page_number,
                lines[idx].strip(),
                0.85,
                block_number,
                _line_confidence(lines, [idx], line_confidences, layout_lines),
                line_number=layout_lines[idx].line_number if layout_lines and idx < len(layout_lines) else idx + 1,
                bbox=_line_bbox([idx], layout_lines),
            )
        elif partial_series:
            result.candidates["series"] = sorted(set(partial_series))
        if len(set(partial_numbers)) == 1:
            observed = partial_numbers[0]
            normalized = normalize_correlativo(observed)
            result.correlativo_original = observed
            result.correlativo_normalized = normalized.normalized
            idx = next(index for series, number, index, _ in signals if number == observed and not series)
            _add_field(
                result,
                "correlativo",
                observed,
                normalized.normalized,
                page_number,
                lines[idx].strip(),
                0.85,
                block_number,
                _line_confidence(lines, [idx], line_confidences, layout_lines),
                line_number=layout_lines[idx].line_number if layout_lines and idx < len(layout_lines) else idx + 1,
                bbox=_line_bbox([idx], layout_lines),
            )
        elif partial_numbers:
            result.candidates["correlativo"] = sorted(set(partial_numbers))
    return unique_pairs


# Pasa una fecha leída a formato ISO. Acepta día/mes/año, año-mes-día y fechas con palabras («14 de Enero del 2019»).
# Si no es una fecha posible, devuelve None.
def _parse_date_str(observed: str) -> str | None:
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(observed.strip(), fmt).date().isoformat()
        except ValueError:
            pass
    m = _SPANISH_TEXT_DATE.search(observed)
    if m:
        day = int(m.group(1))
        month = _SPANISH_MONTHS.get(m.group(2).lower())
        year = int(m.group(3))
        if month:
            try:
                return datetime(year, month, day).date().isoformat()
            except ValueError:
                return None
    return None


# Anota un importe leído. Si ya había otro valor distinto para el mismo concepto, no elige:
# deja ambos como candidatos y el campo vacío, para que se revise.
def _record_amount_field(
    result: ExtractionResult,
    field_name: str,
    observed: str,
    index: int,
    lines: list[str],
    page_number: int,
    block_number: int,
    line_confidences: Mapping[str, float],
    layout_lines: Sequence[LayoutLine] | None,
    currency: str | None,
    *,
    source: str | None = None,
) -> None:
    literal = observed.replace("\u00a0", " ")
    if re.fullmatch(r"[+-]?\s*\d{1,3}(?:\.\d{3})+", literal):
        literal = literal.replace(".", "")
    normalized = _normalize_amount_str(literal)
    if normalized is None:
        return
    existing = getattr(result, field_name)
    line_no = layout_lines[index].line_number if layout_lines and index < len(layout_lines) else index + 1
    bbox = _line_bbox([index], layout_lines)
    score = _line_confidence(lines, [index], line_confidences, layout_lines)
    candidates = set(result.candidates.get(field_name, [])) | {normalized}
    if existing is not None:
        candidates.add(existing)
    conflict = len(candidates) > 1
    _add_field(
        result, field_name, observed.strip(), normalized, page_number, lines[index].strip(),
        0.55 if conflict else 0.90, block_number, score,
        line_number=line_no, bbox=bbox,
        source=source,
        field_status="conflicting_candidates" if conflict else "found",
        decision_comment=f"Moneda del importe: {currency}" if currency else "",
    )
    if conflict:
        result.candidates[field_name] = sorted(candidates)
        setattr(result, field_name, None)
        result.observations.append(f"valores contradictorios para {field_name}")
    else:
        setattr(result, field_name, normalized)


# Recorre las líneas del bloque una sola vez y va sacando lo demás: fechas, nombres del emisor y del receptor,
# moneda, condición de pago, descripción (glosa) e importes.
def _extract_other_fields(
    result: ExtractionResult,
    lines: list[str],
    page_number: int,
    block_number: int,
    line_confidences: Mapping[str, float],
    layout_lines: Sequence[LayoutLine] | None,
) -> None:
    client_start = _client_section_start(lines)
    # Aquí se van juntando las descripciones (glosas) que aparezcan.
    glosa_parts: list[str] = []

    for index, line in enumerate(lines):
        line_no = layout_lines[index].line_number if layout_lines and index < len(layout_lines) else index + 1
        bbox = _line_bbox([index], layout_lines)
        score = _line_confidence(lines, [index], line_confidences, layout_lines)

        # 1. Fecha de emisión.
        if result.issue_date is None:
            date_match = _DATE_LABEL.search(line)
            if date_match and "VENCIMIENTO" not in line[: date_match.start()].upper():
                observed = date_match.group(1)
                normalized = _parse_date_str(observed)
                _add_field(
                    result, "issue_date", observed, normalized, page_number, line.strip(), 0.90, block_number, score, line_number=line_no, bbox=bbox
                )
                result.issue_date = normalized
            elif re.search(r"^\s*(?:FECHA(?:\s+DE)?\s+EMISI[ÓO]N|F\.?\s*EMISI[ÓO]N)\s*[:#-]?\s*$", line, re.IGNORECASE):
                # En representaciones SUNAT de dos columnas el rótulo "Fecha de Emisión" puede quedar
                # separado de su valor ": dd/mm/yyyy" por los rótulos de la misma columna izquierda.
                for offset in range(1, min(6, len(lines) - index)):
                    cand_line = lines[index + offset].strip()
                    if re.search(r"\bVENCIMIENTO\b", cand_line, re.IGNORECASE):
                        break
                    m_next = re.match(r"^[:.\-]*\s*(\d{2}[/.-]\d{2}[/.-]\d{4}|\d{4}-\d{2}-\d{2})\b", cand_line)
                    if m_next:
                        observed = m_next.group(1)
                        normalized = _parse_date_str(observed)
                        if normalized:
                            _add_field(
                                result,
                                "issue_date",
                                observed,
                                normalized,
                                page_number,
                                _excerpt(lines, [index, index + offset], 0),
                                0.88,
                                block_number,
                                score,
                                line_number=line_no,
                                bbox=_line_bbox([index, index + offset], layout_lines),
                            )
                            result.issue_date = normalized
                            break
            else:
                text_date = _SPANISH_TEXT_DATE.search(line)
                if text_date and index <= 14:
                    observed = text_date.group(0)
                    normalized = _parse_date_str(observed)
                    _add_field(
                        result, "issue_date", observed, normalized, page_number, line.strip(), 0.88, block_number, score, line_number=line_no, bbox=bbox
                    )
                    result.issue_date = normalized

        # 2. Fecha de vencimiento.
        if result.due_date is None:
            due_match = _DUE_DATE_LABEL.search(line)
            if due_match:
                observed = due_match.group(1)
                normalized = _parse_date_str(observed)
                _add_field(
                    result, "due_date", observed, normalized, page_number, line.strip(), 0.90, block_number, score, line_number=line_no, bbox=bbox
                )
                result.due_date = normalized
            elif re.search(r"\bFECHA\s+DE\s+VENCIMIENTO\b", line, re.IGNORECASE) and index + 1 < len(lines):
                next_date = _DATE.search(lines[index + 1])
                if next_date:
                    observed = next_date.group(0)
                    normalized = _parse_date_str(observed)
                    _add_field(
                        result, "due_date", observed, normalized, page_number, _excerpt(lines, [index, index + 1], 0), 0.86, block_number, score, line_number=line_no, bbox=bbox
                    )
                    result.due_date = normalized

        # 3. Nombres de las partes (emisor y receptor).
        explicit_issuer = re.search(
            r"\b(?:RAZ[ÓO]N\s+SOCIAL\s+DEL\s+EMISOR|NOMBRE\s+DEL\s+EMISOR)\s*[:#-]\s*(.+)",
            line,
            re.IGNORECASE,
        )
        if explicit_issuer and result.issuer_name is None:
            val = explicit_issuer.group(1).strip()
            result.issuer_name = val
            _add_field(result, "issuer_name", val, val, page_number, line.strip(), 0.92, block_number, score, line_number=line_no, bbox=bbox)

        if result.issuer_name is None and index <= 5 and (client_start is None or index < client_start):
            tiene_forma_societaria = _LEGAL_ENTITY_SUFFIX.search(line) or _LEGAL_ENTITY_SUFFIX_PLAIN.search(line)
            if tiene_forma_societaria and not re.search(r"^\s*(?:RAZ[ÓO]N\s+SOCIAL|SE[ÑN]ORE?\(?S?\)?|CLIENTE)\b", line, re.IGNORECASE):
                cleaned_name = re.sub(r"\bR\.?U\.?C\.?\s*(?:N[°º.]?\s*)?[0-9A-Z]{8,11}.*$", "", line, flags=re.IGNORECASE)
                cleaned_name = re.sub(r"\b(?:FACTURA|BOLETA|NOTA\s+DE\s+CR[EÉ]DITO|NOTA\s+DE\s+D[EÉ]BITO).*$", "", cleaned_name, flags=re.IGNORECASE).strip(" -:")
                if len(cleaned_name) >= 5:
                    result.issuer_name = cleaned_name
                    _add_field(result, "issuer_name", cleaned_name, cleaned_name, page_number, line.strip(), 0.88, block_number, score, line_number=line_no, bbox=bbox)

        recipient_match = re.search(
            r"\b(?:SE[ÑN]ORE?\(?S?\)?|SEFIORE?\(?S?\)?|CLIENTE|ADQUIRENTE|ADQUIRIENTE|RAZ[ÓO]N\s+SOCIAL)\s*[:#-]\s*(.+)",
            line,
            re.IGNORECASE,
        )
        if recipient_match:
            val = re.sub(r"\bR[.:]?U[.:]?C[.:]?\s*[:#-]?\s*[0-9]{11}.*$", "", recipient_match.group(1), flags=re.IGNORECASE).strip()
            if val:
                if client_start is not None and index >= client_start:
                    if result.recipient_name is None:
                        result.recipient_name = val
                        _add_field(result, "recipient_name", val, val, page_number, line.strip(), 0.92, block_number, score, line_number=line_no, bbox=bbox)
                elif re.search(r"\b(?:SE[ÑN]ORE?\(?S?\)?|SEFIORE?\(?S?\)?|CLIENTE|ADQUIRENTE)\b", line, re.IGNORECASE) and result.recipient_name is None:
                    result.recipient_name = val
                    _add_field(result, "recipient_name", val, val, page_number, line.strip(), 0.90, block_number, score, line_number=line_no, bbox=bbox)
                elif result.issuer_name is None and client_start is None:
                    result.issuer_name = val
                    _add_field(result, "issuer_name", val, val, page_number, line.strip(), 0.85, block_number, score, line_number=line_no, bbox=bbox)

        # 4. Moneda.
        # La moneda declarada para una detracción o retención pertenece a ese
        # importe. El comprobante puede estar emitido en una moneda distinta.
        settlement = re.search(r"\b(?:DETRACCI[ÓO]N|RETENCI[ÓO]N|PERCEPCI[ÓO]N)\b", line, re.IGNORECASE)
        curr_match = re.search(
            r"\b(?:TIPO\s+DE\s+)?MONEDA\s*[:#-]\s*[—\-]*\s*(PEN|USD|EUR|SOL(?:ES)?|D[ÓO]LAR(?:ES)?(?:\s+AMERICANO(?:S)?)?)\b",
            line,
            re.IGNORECASE,
        )
        if curr_match and (not settlement or curr_match.start() < settlement.start()):
            observed = curr_match.group(1).strip()
            up = observed.upper()
            normalized = "PEN" if up.startswith("SOL") or up == "PEN" else ("USD" if up.startswith("DOL") or up.startswith("DÓL") or up == "USD" else up)
            result.currency = normalized
            _add_field(result, "currency", observed, normalized, page_number, line.strip(), 0.94, block_number, score, line_number=line_no, bbox=bbox)
        elif result.currency is None and not settlement:
            son_curr = re.search(
                r"\b(?:SON\s*:.*\b|(?:CON|Y)\s+[0-9]{2}/100\s+)(SOLES|NUEVOS\s+SOLES|D[ÓO]LARES(?:\s+AMERICANOS)?)\b",
                line,
                re.IGNORECASE,
            )
            if son_curr:
                observed = son_curr.group(1)
                normalized = "PEN" if "SOL" in observed.upper() else "USD"
                result.currency = normalized
                _add_field(result, "currency", observed, normalized, page_number, line.strip(), 0.90, block_number, score, line_number=line_no, bbox=bbox)
            elif "S/" in line.upper() or "S/." in line.upper():
                result.currency = "PEN"
                _add_field(result, "currency", "S/", "PEN", page_number, line.strip(), 0.75, block_number, score, line_number=line_no, bbox=bbox)
            elif "US$" in line.upper():
                result.currency = "USD"
                _add_field(result, "currency", "US$", "USD", page_number, line.strip(), 0.75, block_number, score, line_number=line_no, bbox=bbox)

        # 5. Condición de pago.
        if result.payment_condition is None:
            pay_match = re.search(r"\b(?:CONDICI[ÓO]N\s+DE\s+PAGO|FORMA\s+DE\s+PAGO)\s*[:#-]\s*(.+)", line, re.IGNORECASE)
            if pay_match:
                val = pay_match.group(1).strip()
                result.payment_condition = val
                _add_field(result, "payment_condition", val, val, page_number, line.strip(), 0.88, block_number, score, line_number=line_no, bbox=bbox)

        # Una descripción con etiqueta propia se conserva tal como aparece.
        # El delimitador evita tomar el encabezado de la tabla como descripción.
        # 6. Descripción: la que viene con su propio rótulo («Descripción:», «Glosa:», «Sustento:»).
        description_match = re.match(r"\s*DESCRIPCI[ÓO]N\s*:\s*(\S.*)", line, re.IGNORECASE)
        if description_match:
            glosa_parts.append(description_match.group(1).strip())

        glosa_match = re.search(r"^\s*(?:GLOSA|SUSTENTO)\s*[:#-]\s*(.+)", line, re.IGNORECASE)
        if glosa_match:
            g_text = glosa_match.group(1).strip(" -")
            following = lines[index + 1:index + 4]
            if following:
                next_line = following[0]
                is_header = re.search(r"^\s*(?:ITEM|UNIDAD|DESCRIPCI[ÓO]N|CODIGO|TOTAL|IGV|IMPORTE)\b", next_line, re.IGNORECASE)
                # Una leyenda tributaria al pie se registra por separado de la descripción.
                is_legend = any(indexes[0] == 0 for _, _, indexes in find_visible_mentions(following))
                if not is_header and not is_legend and not any(amount_mentions(next_line)):
                    g_text = f"{g_text} {next_line.strip()}".strip()
            glosa_parts.append(g_text)

        # 7. Importes rotulados (subtotal, IGV, total...). Los dudosos por el símbolo de moneda se dejan sin leer.
        for field, observed, currency, dudoso in amount_mentions(line):
            if dudoso:
                # El símbolo de moneda se pegó a la cifra; el primer dígito puede no ser
                # suyo. Preferimos dejar el importe en blanco antes que dar por leído un
                # número que quizá no está en el papel.
                result.observations.append(
                    f"importe de «{field}» ilegible: el símbolo de moneda quedó pegado a la cifra "
                    f"en «{line.strip()}»"
                )
                if field not in result.secondary_review_fields:
                    result.secondary_review_fields.append(field)
                continue
            _record_amount_field(
                result, field, observed, index, lines, page_number, block_number,
                line_confidences, layout_lines, currency,
            )

    # Al final, la descripción general es la unión de las glosas encontradas.
    if glosa_parts and not result.general_description:
        result.general_description = " | ".join(dict.fromkeys(glosa_parts))
        _add_field(
            result,
            "general_description",
            result.general_description,
            result.general_description,
            page_number,
            result.general_description,
            0.90,
            block_number,
        )


# Moldes para las filas del detalle de productos o servicios.
# Éste reconoce filas con número de ítem, descripción, unidad, cantidad, precio y total.
_ITEM_ROW_TABULAR = re.compile(
    r"^\s*(?P<idx>[0-9]{1,4})\s+(?P<desc>.+?)\s+(?P<unit>ZZ|NIU|KWH|KW|MWH|UND|UN|GLB|SER|MES|DIA|HOR|PA|CJ|BOL|KG|TON|LT|GAL|M3|M2|MT|PZA)\s+"
    r"(?P<qty>[0-9][0-9,.]*)\s+(?P<uprice>[0-9][0-9,.]*)"
    r"(?:\s+(?P<punit>[0-9][0-9,.]*))?\s+(?P<total>[0-9][0-9,.]*)\s*$",
    re.IGNORECASE,
)

# Fila de detalle con sólo descripción y total (sin cantidad ni unidad), con los miles separados por espacio
# y los decimales con coma.
_ITEM_ROW_ELECTRO = re.compile(
    r"^\s*(?P<desc>[A-Za-zÁÉÍÓÚáéíóúÑñ][A-Za-z0-9ÁÉÍÓÚáéíóúÑñ\s().-]{6,}?)\s+(?P<total>[0-9]{1,3}(?:\s+[0-9]{3})+,[0-9]{2}|[0-9]{1,6},[0-9]{2})\s*$"
)

# Cómo reconocer el encabezado de la tabla de detalle y dónde termina (subtotal, total, IGV...).
_DETAIL_HEADER = re.compile(r"\b(?:DESCRIPCI[ÓO]N|PRODUCTO)\b", re.IGNORECASE)
_DETAIL_COLUMN = re.compile(r"\b(?:CANT(?:IDAD)?|UNITARIO|P\.?\s*UNIT|UNIDAD|[ÍI]TEM|C[ÓO]DIGO)\b", re.IGNORECASE)
_DETAIL_END = re.compile(
    r"^(?:SUB\s*TOTAL|VALOR\s+(?:DE\s+)?VENTA|IMPORTE\s+TOTAL|TOTAL\b|IGV\b|ISC\b|ICBPER\b|SON\s*:|ANTICIPOS?\b|OP\.\s*GRAVADA|DETRACCI[ÓO]N\b)",
    re.IGNORECASE,
)
# Piezas para armar los moldes de filas de tabla: un número, una unidad de medida y tres formas de fila con descripción.
_TABLE_AMOUNT = r"[0-9]+(?:[,.][0-9]+)*"
_TABLE_UNIT = r"(?:UNIDAD(?:ES)?|NIU|UND|UN|ZZ|KWH|KW|GLB|SER|MES|DIA|HOR|KG|LTR|M[23])"
_TABLE_DESCRIPTION = (
    re.compile(rf"^{_TABLE_AMOUNT}\s+(?:[-–]\s+)?(?P<desc>.+?)\s+{_TABLE_UNIT}(?:\s+{_TABLE_AMOUNT}){{2,4}}$", re.IGNORECASE),
    re.compile(rf"^{_TABLE_AMOUNT}\s+{_TABLE_UNIT}\s+(?P<desc>.+?)(?:\s+{_TABLE_AMOUNT}){{2,4}}$", re.IGNORECASE),
    re.compile(rf"^(?P<desc>[A-ZÁÉÍÓÚÑ].+?)\s+\d+(?:[,.]\d+|\s+\d{{2}})?\s+{_TABLE_AMOUNT}\s+{_TABLE_AMOUNT}$", re.IGNORECASE),
)


# Recorre el bloque buscando la tabla de detalle y saca cada ítem (descripción, cantidad, unidad, precio y total).
# Con ellos arma la descripción general. Sólo se guardan filas que encajan en un molde; nada se completa por suposición.
def _extract_line_items(
    result: ExtractionResult,
    lines: list[str],
    page_number: int,
    block_number: int,
    line_confidences: Mapping[str, float],
    layout_lines: Sequence[LayoutLine] | None,
) -> None:
    # Estado: ¿estamos dentro de la tabla de detalle? Más los datos para unir descripciones partidas en varias líneas.
    in_electro_table = False
    descriptions: list[tuple[str, int, int]] = []
    has_code_column = False
    previous_detail = None

    for index, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            continue
        # La tabla empieza cuando una línea con «Descripción» o «Producto» va junto a nombres de columna.
        header = " ".join(lines[max(0, index - 1):index + 2])
        if _DETAIL_HEADER.search(stripped) and _DETAIL_COLUMN.search(header):
            in_electro_table = True
            has_code_column = bool(re.search(r"\bC[ÓO]DIGO\b", header, re.IGNORECASE))
            continue
        # La tabla termina al llegar a los totales o a cualquier importe rotulado.
        if _DETAIL_END.search(stripped) or any(amount_mentions(stripped)):
            in_electro_table = False
            previous_detail = None
            continue

        # Fila completa con ítem, unidad, cantidad, precio y total: se guarda como un ítem.
        m_item = _ITEM_ROW_TABULAR.match(stripped)
        if m_item:
            raw_idx = m_item.group("idx")
            desc = m_item.group("desc").strip(" .|")
            unit = m_item.group("unit").upper()
            qty_obs = m_item.group("qty")
            uprice_obs = m_item.group("uprice")
            total_obs = m_item.group("total")
            line_no = layout_lines[index].line_number if layout_lines and index < len(layout_lines) else index + 1
            bbox = _line_bbox([index], layout_lines)
            ocr_conf = _line_confidence(lines, [index], line_confidences, layout_lines)
            result.line_items.append(
                LineItem(
                    item_index=len(result.line_items) + 1,
                    code=raw_idx if len(raw_idx) >= 3 else None,
                    description_observed=desc,
                    description_normalized=re.sub(r"\s+", " ", desc),
                    quantity_observed=qty_obs,
                    quantity_normalized=_normalize_quantity_str(qty_obs),
                    unit_observed=unit,
                    unit_normalized=unit,
                    unit_price_observed=uprice_obs,
                    unit_price_normalized=_normalize_amount_str(uprice_obs),
                    line_total_observed=total_obs,
                    line_total_normalized=_normalize_amount_str(total_obs),
                    page_number=page_number,
                    block_number=block_number,
                    line_number=line_no,
                    bbox=bbox,
                    excerpt=stripped,
                    source=result.reading_source,
                    ocr_confidence=ocr_conf,
                    heuristic_confidence=0.93,
                )
            )
            descriptions.append((desc, index, index))
            previous_detail = None
            continue

        # Dentro de la tabla: filas con sólo descripción y total, o descripciones sueltas que luego se unen.
        if in_electro_table:
            m_elec = _ITEM_ROW_ELECTRO.match(stripped)
            if m_elec:
                desc = m_elec.group("desc").strip()
                if index + 1 < len(lines) and len(lines[index + 1].strip()) <= 28 and not re.search(r"[0-9:]", lines[index + 1]):
                    desc = f"{desc} {lines[index + 1].strip()}".strip()
                total_obs = m_elec.group("total")
                total_norm = _normalize_amount_str(total_obs)
                line_no = layout_lines[index].line_number if layout_lines and index < len(layout_lines) else index + 1
                bbox = _line_bbox([index], layout_lines)
                ocr_conf = _line_confidence(lines, [index], line_confidences, layout_lines)
                result.line_items.append(
                    LineItem(
                        item_index=len(result.line_items) + 1,
                        code=None,
                        description_observed=desc,
                        description_normalized=re.sub(r"\s+", " ", desc),
                        quantity_observed=None,
                        quantity_normalized=None,
                        unit_observed=None,
                        unit_normalized=None,
                        unit_price_observed=total_obs,
                        unit_price_normalized=total_norm,
                        line_total_observed=total_obs,
                        line_total_normalized=total_norm,
                        page_number=page_number,
                        block_number=block_number,
                        line_number=line_no,
                        bbox=bbox,
                        excerpt=stripped,
                        source=result.reading_source,
                        ocr_confidence=ocr_conf,
                        heuristic_confidence=0.90,
                        missing_reason="quantity_and_unit_not_present_in_source_table",
                    )
                )
                descriptions.append((desc, index, index))
                previous_detail = None
                continue
            # Otras formas de fila: de ellas se guarda sólo la descripción.
            for pattern in _TABLE_DESCRIPTION:
                match = pattern.fullmatch(stripped)
                if not match:
                    continue
                desc = match["desc"].strip(" .|")
                words = desc.split(maxsplit=2)
                # En algunas tablas el OCR junta la columna de código con la descripción.
                if has_code_column and len(words) == 3 and words[0].casefold() == words[1].casefold():
                    desc = f"{words[1]} {words[2]}"
                descriptions.append((desc, index, index))
                previous_detail = len(descriptions) - 1
                break
            else:
                if previous_detail is not None:
                    desc, start, end = descriptions[previous_detail]
                    if end == index - 1 and desc.endswith("-") and re.match(r"^[A-ZÁÉÍÓÚÑ]", stripped, re.IGNORECASE):
                        descriptions[previous_detail] = (f"{desc} {stripped}", start, index)
                previous_detail = None

    # La descripción general es la unión de todas las descripciones encontradas.
    if descriptions and not result.general_description:
        result.general_description = " | ".join(dict.fromkeys(desc for desc, _, _ in descriptions))
        for desc, start, end in descriptions:
            indexes = list(range(start, end + 1))
            _add_field(
                result, "general_description", desc, desc, page_number,
                _excerpt(lines, indexes, radius=0), 0.88, block_number,
                _line_confidence(lines, indexes, line_confidences, layout_lines),
                line_number=layout_lines[start].line_number if layout_lines and start < len(layout_lines) else start + 1,
                bbox=_line_bbox(indexes, layout_lines),
            )


# Controles que se calculan al final, con el comprobante ya completo: falta de descripción, de fecha o de razón social
# del emisor, lecturas de OCR dudosas, cuadre de importes y coherencia del detalle.
# Cada control que falla se anota para la hoja revision.
def _compute_derived_checks(result: ExtractionResult, minimum_confidence: float) -> None:
    if result.detection_status != DETECTED:
        return
    if not result.general_description:
        result.derived_checks.append({
            "check": "missing_description",
            "consistent": False,
            "field": "Descripción",
            "page": result.page_start,
            "excerpt": result.detection_excerpt or "",
            "reason": "Revisa la glosa o las filas de productos y servicios en el original.",
        })
        result.secondary_review_fields.append("general_description")
    # Un comprobante detectado sin fecha de emisión legible tiene que quedar anotado:
    # la celda vacía del Excel no explica por sí sola de dónde viene el vacío.
    if not result.issue_date:
        result.derived_checks.append({
            "check": "missing_issue_date",
            "consistent": False,
            "field": "Fecha de emisión",
            "page": result.page_start,
            "excerpt": result.detection_excerpt or "",
            "reason": "No se leyó la fecha de emisión; búscala en el original y anótala a mano.",
        })
        result.secondary_review_fields.append("issue_date")
    # La razón social del emisor en blanco tampoco puede quedar muda: el RUC por sí solo
    # no la rellena y el lector no la adivina a partir de la posición en la página.
    if not result.issuer_name:
        result.derived_checks.append({
            "check": "missing_issuer_name",
            "consistent": False,
            "field": "Razón social del emisor",
            "page": result.page_start,
            "excerpt": result.detection_excerpt or "",
            "reason": "No se leyó la razón social del emisor; compárala con el original o con el padrón.",
        })
        result.secondary_review_fields.append("issuer_name")
    # Este umbral indica qué lecturas conviene mirar de nuevo. No cambia el tipo
    # detectado ni reemplaza los caracteres que leyó Tesseract.
    fields = {
        "document_type": ("Denominación", result.document_type),
        "emitter_ruc": ("RUC emisor", result.emitter_ruc),
        "recipient_ruc": ("RUC receptor", result.recipient_ruc),
        "series": ("Serie", result.series),
        "correlativo": ("Correlativo", result.correlativo_normalized),
    }
    for field, (label, value) in fields.items():
        evidence = [item for item in result.field_evidence.get(field, []) if item.normalized_value == value]
        if not value or not evidence or any(item.ocr_confidence is None for item in evidence):
            continue
        best = max(evidence, key=lambda item: item.ocr_confidence)
        if best.ocr_confidence >= minimum_confidence:
            continue
        result.derived_checks.append({
            "check": "low_ocr_confidence",
            "consistent": False,
            "field": label,
            "observed_value": best.observed_value,
            "page": best.page_number,
            "excerpt": best.excerpt,
            "reason": f"La lectura obtuvo {best.ocr_confidence * 100:.1f} de 100. El umbral es {minimum_confidence * 100:.0f}.",
        })
        if field not in result.secondary_review_fields:
            result.secondary_review_fields.append(field)
    # Se compara el total una vez, cuando ya se reunieron las páginas del comprobante.
    check = reconcile_amounts(result)
    result.derived_checks.append(check)
    if check["status"] != "cuadra":
        result.secondary_review_fields.append("amounts")
        result.observations.append(check.get("reason") or (
            f"Los componentes suman {check['expected_total']} y el total leído es {check['observed_total']}."
        ))

    # Verificar conciliación secundaria de ítems cuando hay cantidad, valor unitario e importe de línea
    item_discrepancies: list[str] = []
    for item in result.line_items:
        for obs_label, obs_val in (("unitario", item.unit_price_observed), ("total", item.line_total_observed)):
            if obs_val and re.fullmatch(r"-?\d{1,3}(?:,\d{3})+,\d{2}", obs_val.strip()):
                item_discrepancies.append(
                    f"ítem {item.item_index} (pág. {item.page_number}, lín. {item.line_number or '?'}): "
                    f"puntuación decimal dudosa por OCR en {obs_label}={obs_val}"
                )
        if item.quantity_normalized and item.unit_price_normalized and item.line_total_normalized:
            try:
                q_val = Decimal(item.quantity_normalized)
                u_val = Decimal(item.unit_price_normalized)
                t_val = Decimal(item.line_total_normalized)
                expected_line = (q_val * u_val).quantize(Decimal("0.01"))
                if abs(expected_line - t_val) > Decimal("0.10"):
                    item_discrepancies.append(
                        f"ítem {item.item_index} (pág. {item.page_number}, lín. {item.line_number or '?'}): "
                        f"cantidad={item.quantity_observed} × unitario={item.unit_price_observed} != total={item.line_total_observed}"
                    )
            except InvalidOperation:
                item_discrepancies.append(
                    f"ítem {item.item_index} (pág. {item.page_number}): formato numérico dudoso en detalle"
                )
    if item_discrepancies:
        result.derived_checks.append(
            {
                "check": "line_item_arithmetic_reconciliation",
                "consistent": False,
                "discrepancies": item_discrepancies,
            }
        )
        result.observations.append(
            f"revisión secundaria de detalle en {len(item_discrepancies)} ítem(s): {'; '.join(item_discrepancies)}"
        )
        if "line_items" not in result.secondary_review_fields:
            result.secondary_review_fields.append("line_items")


# A los campos que no se encontraron se les pone un estado explícito («missing_in_source» o «not_applicable_non_receipt»)
# en lugar de dejarlos sin estado.
def _populate_missing_field_statuses(result: ExtractionResult) -> None:
    expected_fields = (
        "emitter_ruc",
        "recipient_ruc",
        "document_type",
        "series",
        "correlativo",
        "issue_date",
        "due_date",
        "issuer_name",
        "recipient_name",
        "currency",
        *AMOUNT_FIELDS,
        "general_description",
    )
    for fname in expected_fields:
        if fname not in result.field_statuses:
            result.field_statuses[fname] = (
                "not_applicable_non_receipt" if result.detection_status == NOT_RECEIPT else "missing_in_source"
            )


# Decide si la identidad del comprobante está completa (tipo, RUC emisor válido, serie y correlativo), incompleta o en conflicto,
# y deja escrito el motivo. También calcula la clave fiscal y si el documento se podría renombrar.
def _finalize_identity_and_reasons(
    result: ExtractionResult,
    catalog: Mapping[str, str],
    minimum_confidence: float,
    *,
    conflicting_types: bool = False,
    conflicting_serials: bool = False,
) -> None:
    """Evalúa el estado de identidad de forma independiente al estado de detección."""
    result.comprobante_detectado = result.detection_status == DETECTED
    if not result.comprobante_detectado:
        result.identity_status = IDENTITY_NOT_APPLICABLE
        result.identidad_completa = False
        result.renombrable = False
        result.identity_reason = "No aplica estado de identidad fiscal porque el bloque no es un CDP."
        result.identity_excerpt = None
        if not result.detection_reason:
            result.detection_reason = "; ".join(result.observations) or result.classification_rule
        return

    # Paso 1: ¿cada pieza de la identidad es válida?
    ruc_validation = validate_ruc(result.emitter_ruc or "")
    has_valid_emitter = bool(ruc_validation.format_valid and ruc_validation.checksum_valid)
    has_valid_type = result.document_type in catalog
    has_valid_series = bool(result.series and re.fullmatch(r"[A-Za-z0-9]{1,20}", result.series))
    has_valid_corr = bool(result.correlativo_normalized and re.fullmatch(r"[0-9]{8}", result.correlativo_normalized))

    conflicting_emitters = len(result.candidates.get("emitter_ruc", [])) > 1
    has_doubtful_chars = any(
        "dudoso" in obs.lower() or "ambigu" in obs.lower() or "conflicto" in obs.lower()
        for obs in result.observations
    )

    # Paso 2: qué piezas faltan o necesitan verificación.
    missing_fields: list[str] = []
    if not has_valid_type:
        missing_fields.append("document_type")
    if not has_valid_emitter:
        missing_fields.append("emitter_ruc")
    if not has_valid_series:
        missing_fields.append("series")
    elif "series" in result.secondary_review_fields:
        missing_fields.append("series")
    if not has_valid_corr:
        missing_fields.append("correlativo")

    # Paso 3: el estado de la identidad: en conflicto (candidatos múltiples), incompleta (faltan piezas o hay caracteres dudosos)
    # o completa.
    if conflicting_types or conflicting_serials or conflicting_emitters:
        result.identity_status = IDENTITY_CONFLICT
        result.identidad_completa = False
        conflict_parts: list[str] = []
        if conflicting_types:
            conflict_parts.append(f"tipos={result.candidates.get('document_type', [])}")
        if conflicting_serials:
            conflict_parts.append(
                f"series={result.candidates.get('series', [])} / correlativos={result.candidates.get('correlativo', [])}"
            )
        if conflicting_emitters:
            conflict_parts.append(f"emisores={result.candidates.get('emitter_ruc', [])}")
        result.identity_reason = f"Identidad en conflicto por candidatos múltiples: {', '.join(conflict_parts)}."
    elif missing_fields or has_doubtful_chars:
        result.identity_status = IDENTITY_INCOMPLETE
        result.identidad_completa = False
        if missing_fields:
            result.identity_reason = f"Identidad incompleta; faltan o requieren verificación: {', '.join(missing_fields)}."
        else:
            result.identity_reason = "Identidad incompleta por caracteres dudosos (O/0 o I/1) que requieren revisión."
    else:
        result.identity_status = IDENTITY_COMPLETE
        result.identidad_completa = True
        result.identity_reason = (
            f"Identidad fiscal completa: tipo={result.document_type}, "
            f"emisor={result.emitter_ruc}, serie={result.series}, correlativo={result.correlativo_normalized}."
        )

    # Paso 4: fragmento de sustento de la identidad y de la detección, a partir de la evidencia observada.
    identity_excerpts: list[str] = []
    for fname in ("document_type", "emitter_ruc", "series", "correlativo"):
        ev_list = result.field_evidence.get(fname) or []
        if ev_list and ev_list[0].excerpt:
            if ev_list[0].excerpt not in identity_excerpts:
                identity_excerpts.append(ev_list[0].excerpt)
    result.identity_excerpt = " | ".join(identity_excerpts) if identity_excerpts else None
    if not result.detection_excerpt:
        result.detection_excerpt = result.identity_excerpt or (
            result.extracted_text.splitlines()[0].strip() if result.extracted_text else None
        )
    if not result.detection_reason:
        result.detection_reason = "; ".join(result.observations[:2]) or result.classification_rule

    # Paso 5: la clave fiscal sólo existe si la identidad está completa.
    if result.identity_key is not None:
        result.candidate_fiscal_key = (
            f"{result.emitter_ruc}-{result.document_type}-{(result.series or '').upper()}-{result.correlativo_normalized}"
        )
    else:
        result.candidate_fiscal_key = None
    min_confidence = min(
        (result.confidence.get(name, 0.0) for name in ("emitter_ruc", "document_type", "series", "correlativo")),
        default=0.0,
    )
    result.renombrable = bool(result.identidad_completa and not result.read_error and min_confidence >= minimum_confidence)


# Resultado para una página de XML: se excluye por formato y no cuenta como comprobante.
def _excluded_xml_result(page_number: int) -> ExtractionResult:
    # Los archivos XML se excluyen del flujo de detección visual de CDP.
    result = ExtractionResult(
        page_start=page_number,
        page_end=page_number,
        block_start=1,
        block_end=1,
        reading_source="excluded_xml",
        extracted_text="",
        detection_status=NOT_RECEIPT,
        identity_status=IDENTITY_NOT_APPLICABLE,
        comprobante_detectado=False,
        non_receipt_type="xml_excluido_por_formato",
        classification_rule="excluded_xml_format",
        detection_reason="Archivo XML excluido por alcance del flujo de comprobantes visibles.",
    )
    result.observations.append("archivo XML excluido por formato; no alimenta identificación ni conteo de CDP")
    _populate_missing_field_statuses(result)
    result.decision_comment = "Archivo XML excluido por formato del flujo de detección visual."
    return result


# Analiza un bloque de una página y decide qué es: comprobante, no comprobante, indeterminado o error de lectura.
# Extrae todos sus datos y deja escrito por qué decidió así.
def _parse_block(
    lines: list[str],
    page_number: int,
    block_number: int,
    page_status: str,
    line_confidences: Mapping[str, float],
    catalog: Mapping[str, str],
    minimum_confidence: float,
    *,
    reading_source: str = "native_pdf_text",
    layout_lines: Sequence[LayoutLine] | None = None,
    target_recipient_ruc: str | None = None,
    page_error_message: str | None = None,
) -> ExtractionResult:
    # Paso 1: preparar el resultado y atender los casos que no necesitan análisis:
    # XML excluido, error de lectura o página vacía.
    result = ExtractionResult(
        page_start=page_number,
        page_end=page_number,
        block_start=block_number,
        block_end=block_number,
        reading_source=reading_source,
    )
    if reading_source == "excluded_xml" or page_status == "excluded_by_format":
        result.detection_status = NOT_RECEIPT
        result.identity_status = IDENTITY_NOT_APPLICABLE
        result.comprobante_detectado = False
        result.non_receipt_type = "xml_excluido_por_formato"
        result.classification_rule = "excluded_xml_format"
        result.detection_reason = "Archivo XML excluido por formato del flujo de detección visual."
        result.observations.append("archivo XML excluido por formato")
        result.decision_comment = result.detection_reason
        _populate_missing_field_statuses(result)
        return result

    result.extracted_text = "\n".join(lines).strip()
    result.read_error = page_status in _READ_ERRORS
    if result.read_error:
        detail = f"página {page_number}: {page_status}"
        if page_error_message:
            detail = f"{detail} ({page_error_message})"
        result.read_error_details.append(detail)
    nonempty = [line.strip() for line in lines if line.strip()]
    if not nonempty:
        if result.read_error:
            result.detection_status = READ_ERROR
            result.classification_rule = "read_error_unreadable_page"
            result.detection_reason = "; ".join(result.read_error_details) or f"Error de lectura en página {page_number}."
            result.observations.append(result.detection_reason)
            result.decision_comment = f"Error de lectura en página {page_number}: {page_status}."
        else:
            result.detection_status = AMBIGUOUS
            result.classification_rule = "empty_or_unreadable_page"
            result.detection_reason = "Página sin texto legible tras lectura nativa y/u OCR."
            result.observations.append("página sin texto legible")
            result.decision_comment = "Página sin texto extraíble o en blanco."
        result.comprobante_detectado = False
        result.identity_status = IDENTITY_NOT_APPLICABLE
        _populate_missing_field_statuses(result)
        return result

    # Paso 2: ¿es un documento de soporte conocido (HES, orden de compra, estado de cuenta...)? Se descarta antes de buscar datos fiscales.
    non_receipt_type, non_receipt_rule = _detect_non_receipt_type(nonempty)
    if non_receipt_type is not None:
        result.detection_status = NOT_RECEIPT
        result.identity_status = IDENTITY_NOT_APPLICABLE
        result.comprobante_detectado = False
        result.non_receipt_type = non_receipt_type
        result.classification_rule = non_receipt_rule or "non_receipt"
        result.detection_reason = f"Documento de soporte o no fiscal clasificado como {non_receipt_type}."
        result.detection_excerpt = nonempty[0][:160]
        result.observations.append(f"documento clasificado como no_cdp ({non_receipt_type})")
        result.decision_comment = f"No es comprobante de pago: se identificó estructura de {non_receipt_type}."
        _populate_missing_field_statuses(result)
        return result

    # Paso 3: extraer, en este orden, el tipo, los RUC, la serie y el correlativo, los demás datos,
    # el detalle y las leyendas tributarias visibles.
    type_codes = _extract_types(result, nonempty, catalog, page_number, block_number, line_confidences, layout_lines)
    ruc_seen = _extract_rucs(result, nonempty, page_number, block_number, line_confidences, layout_lines, target_recipient_ruc)
    serial_pairs = _extract_serials(result, nonempty, page_number, block_number, line_confidences, layout_lines)
    _extract_other_fields(result, nonempty, page_number, block_number, line_confidences, layout_lines)
    _extract_line_items(result, nonempty, page_number, block_number, line_confidences, layout_lines)
    for field, code, indexes in find_visible_mentions(nonempty):
        excerpt = _excerpt(nonempty, indexes, radius=0)
        _add_field(
            result, field, excerpt, code, page_number, excerpt, 0.90, block_number,
            _line_confidence(nonempty, indexes, line_confidences, layout_lines),
            bbox=_line_bbox(indexes, layout_lines),
            decision_comment="Leyenda visible en el documento; no acredita pago ni depósito.",
        )

    # Paso 4: decidir. Hay comprobante si el título indica un único tipo y hay al menos un dato de respaldo
    # (RUC, número, importe o fecha).
    has_generic_receipt = bool(_GENERIC_RECEIPT.search("\n".join(nonempty)))
    has_amount = any(getattr(result, field) is not None for field in AMOUNT_FIELDS)
    has_date = bool(result.issue_date or result.due_date)
    has_partial_serial = bool(result.series or result.correlativo_original)
    has_fiscal_signal = bool(type_codes or ruc_seen or serial_pairs or has_partial_serial or has_amount or has_date or has_generic_receipt)
    conflicting_types = len(type_codes) > 1
    conflicting_serials = len(serial_pairs) > 1
    support = bool(serial_pairs or has_partial_serial or ruc_seen or has_amount or has_date)

    # Un RUC, un número o un monto ayudan a identificar el documento cuando el título
    # ya indica su tipo. Esos datos por sí solos también aparecen en soportes y anexos.
    if len(type_codes) == 1 and support:
        result.detection_status = DETECTED
        result.classification_rule = "tipo_documental_y_soporte"
        result.detection_reason = "Denominación del documento en el encabezado y señales de respaldo en el mismo bloque."
    elif not type_codes and _AIR_PURCHASE.search(result.extracted_text) and _AIR_CONTEXT.search(result.extracted_text):
        result.detection_status = AMBIGUOUS
        result.non_receipt_type = "documento_aereo_por_tipificar"
        result.classification_rule = "compra_aerea_sin_denominacion_fiscal"
        result.detection_reason = "Comprobante de compra aérea sin denominación fiscal admitida; requiere revisión como documento aéreo."
    elif result.read_error and not has_fiscal_signal:
        result.detection_status = READ_ERROR
        result.classification_rule = "read_error"
        result.detection_reason = "; ".join(result.read_error_details) or "Error de lectura en la página."
    elif has_fiscal_signal:
        result.detection_status = AMBIGUOUS
        result.classification_rule = "ambiguous_partial_fiscal_signals"
        result.detection_reason = "Existen menciones aisladas (RUC, monto o numeración), pero sin estructura suficiente de comprobante."
        result.observations.append("hay señales fiscales aisladas, pero no evidencia suficiente para confirmar un CDP")
    else:
        result.detection_status = NOT_RECEIPT
        result.classification_rule = "non_receipt_no_fiscal_signals"
        result.detection_reason = "Contenido legible sin señales de comprobante de pago."
        result.observations.append("contenido legible sin señales de comprobante de pago")

    # Paso 5: con el tipo ya decidido, evaluar si la identidad está completa y armar el comentario final de la decisión.
    _finalize_identity_and_reasons(
        result,
        catalog,
        minimum_confidence,
        conflicting_types=conflicting_types,
        conflicting_serials=conflicting_serials,
    )
    _populate_missing_field_statuses(result)

    if result.comprobante_detectado:
        result.decision_comment = (
            f"Comprobante detectado ({result.document_type or '?'} {result.series or '?'}-{result.correlativo_normalized or result.correlativo_original or '?'}); "
            f"identidad={result.identity_status}, emisor={result.emitter_ruc or 'pendiente'}, "
            f"receptor={result.recipient_ruc or 'pendiente'} (match={result.recipient_ruc_match})."
        )
    elif result.detection_status == READ_ERROR:
        result.decision_comment = result.detection_reason or "Error de lectura."
    elif result.detection_status == AMBIGUOUS:
        result.decision_comment = "; ".join(result.observations) or "Evidencia insuficiente o indeterminada."
    else:
        result.decision_comment = "; ".join(result.observations) or "Sin señales de comprobante de pago."
    return result


# Acepta una página en dos formas (la tupla sencilla o el objeto PageRead) y devuelve sus datos en un formato único.
def _page_values(page: Any) -> tuple[int, str, str, Mapping[str, float], str, Sequence[LayoutLine], Mapping[str, Any] | None, str | None]:
    if isinstance(page, tuple) and len(page) >= 2:
        number, text = int(page[0]), str(page[1])
        status = str(page[2]) if len(page) > 2 else "read"
        line_confidences = page[3] if len(page) > 3 else {}
        return number, text, status, line_confidences or {}, "native_pdf_text", [], None, None
    return (
        int(getattr(page, "page_number", 1)),
        str(getattr(page, "text", "")),
        str(getattr(page, "status", "read")),
        getattr(page, "line_confidences", {}) or {},
        str(getattr(page, "source", "native_pdf_text")),
        getattr(page, "layout_lines", None) or [],
        getattr(page, "structured_data", None),
        getattr(page, "error_message", None),
    )


# ¿Dos resultados tienen la misma identidad completa (emisor, tipo, serie y correlativo)?
# Sirve para unir páginas que continúan el mismo comprobante.
def _same_complete_identity(left: ExtractionResult, right: ExtractionResult) -> bool:
    if left.identity_key is not None and left.identity_key == right.identity_key:
        return True
    if (
        left.document_type
        and left.document_type == right.document_type
        and left.series
        and right.series
        and left.series.casefold() == right.series.casefold()
        and (left.correlativo_normalized or left.correlativo_original)
        and (left.correlativo_normalized or left.correlativo_original)
        == (right.correlativo_normalized or right.correlativo_original)
    ):
        return True
    return False


# Une una página de continuación a la del comprobante: suma evidencias, completa los campos que faltaban,
# conserva ambos valores si un importe cambia entre páginas, evita repetir ítems y recalcula la identidad.
def _merge_page(
    target: ExtractionResult,
    continuation: ExtractionResult,
    catalog: Mapping[str, str] | None = None,
    minimum_confidence: float = 0.70,
) -> None:
    target.page_end = max(target.page_end, continuation.page_end)
    target.block_end = continuation.block_end
    if continuation.extracted_text:
        target.extracted_text = f"{target.extracted_text}\n\n{continuation.extracted_text}".strip()
    for name, items in continuation.field_evidence.items():
        target.field_evidence.setdefault(name, []).extend(items)
    for name, items in continuation.evidence.items():
        target.evidence.setdefault(name, []).extend(items)
    for name, value in continuation.confidence.items():
        target.confidence[name] = min(target.confidence.get(name, 1.0), value)
    for name, values in continuation.candidates.items():
        target.candidates[name] = sorted(set(target.candidates.get(name, [])) | set(values))

    # Una repetición en otra página no suma de nuevo. Si el importe cambia,
    # guardamos ambos valores para que el cuadre quede en revisión.
    for field in AMOUNT_FIELDS:
        values = set(target.candidates.get(field, []))
        values.update(value for value in (getattr(target, field), getattr(continuation, field)) if value is not None)
        if len(values) > 1:
            target.candidates[field] = sorted(values)
            setattr(target, field, None)
        elif values:
            setattr(target, field, next(iter(values)))

    for field_name in (
        "emitter_ruc",
        "recipient_ruc",
        "issuer_name",
        "recipient_name",
        "issue_date",
        "due_date",
        "currency",
        "payment_condition",
        "referenced_document",
        "general_description",
    ):
        if getattr(target, field_name) is None and getattr(continuation, field_name) is not None:
            setattr(target, field_name, getattr(continuation, field_name))
            target.field_statuses[field_name] = continuation.field_statuses.get(field_name, "found")
            if field_name == "recipient_ruc":
                target.recipient_ruc_match = continuation.recipient_ruc_match
                target.recipient_role_status = continuation.recipient_role_status
            elif field_name == "emitter_ruc":
                target.issuer_role_status = continuation.issuer_role_status

    for tp_ruc in continuation.third_party_rucs:
        if tp_ruc not in target.third_party_rucs:
            target.third_party_rucs.append(tp_ruc)

    existing_descs = {
        (it.description_normalized or it.description_observed, it.line_total_normalized or it.line_total_observed)
        for it in target.line_items
    }
    for item in continuation.line_items:
        sig = (item.description_normalized or item.description_observed, item.line_total_normalized or item.line_total_observed)
        if sig not in existing_descs:
            item.item_index = len(target.line_items) + 1
            target.line_items.append(item)
            existing_descs.add(sig)

    target.observations.extend(item for item in continuation.observations if item not in target.observations)
    target.read_error = target.read_error or continuation.read_error
    target.read_error_details.extend(continuation.read_error_details)
    if target.comprobante_detectado:
        _finalize_identity_and_reasons(
            target,
            catalog or DEFAULT_DOCUMENT_TYPES,
            minimum_confidence,
            conflicting_types=len(target.candidates.get("document_type", [])) > 1,
            conflicting_serials=len(target.candidates.get("series", [])) > 1,
        )
    target.renombrable = target.identidad_completa and not target.read_error


# Punto de entrada de este módulo: recibe las páginas leídas de un documento y devuelve los resultados
# (uno por comprobante o por bloque). Sólo usa la evidencia de esas páginas; nunca consulta otros archivos ni corridas anteriores.
def parse_document(
    pages: Sequence[Any],
    catalog: Mapping[str, str] | None = None,
    minimum_confidence: float = 0.70,
    *,
    target_recipient_ruc: str | None = None,
) -> list[ExtractionResult]:
    """Detecta los grupos de comprobantes a partir de la evidencia de cada página; nunca consulta otros archivos ni corridas anteriores."""
    # Paso 1: leer cada página y partirla en bloques (puede haber más de un comprobante, o un soporte junto a un comprobante).
    type_catalog = dict(catalog or DEFAULT_DOCUMENT_TYPES)
    page_results: list[ExtractionResult] = []
    for page in sorted((_page_values(item) for item in pages), key=lambda item: item[0]):
        page_number, text, page_status, line_confidences, reading_source, layout_lines, structured_data, page_error_message = page
        if reading_source in {"xml_structured", "excluded_xml"} or page_status == "excluded_by_format":
            page_results.append(_excluded_xml_result(page_number))
            continue

        lines = [line.strip() for line in text.splitlines()]
        blocks = _split_blocks(lines, type_catalog)
        use_layout = layout_lines if len(blocks) == 1 else None
        for block_number, block in enumerate(blocks, start=1):
            page_results.append(
                _parse_block(
                    block,
                    page_number,
                    block_number,
                    page_status,
                    line_confidences,
                    type_catalog,
                    minimum_confidence,
                    reading_source=reading_source,
                    layout_lines=use_layout,
                    target_recipient_ruc=target_recipient_ruc,
                    page_error_message=page_error_message,
                )
            )

    # Paso 2: unir las páginas que continúan a la anterior (marcas de «página 2 de 3», anexos de factura negociable
    # o la misma identidad).
    grouped: list[ExtractionResult] = []
    for current in page_results:
        previous = grouped[-1] if grouped else None
        adjacent = previous is not None and current.page_start == previous.page_end + 1
        continuation_marker = bool(
            _PAGE_MARKER.search(current.extracted_text)
            or _CONTINUATION.search(current.extracted_text)
            or current.non_receipt_type == "anexo_factura_negociable"
        )
        if adjacent and previous is not None:
            if previous.detection_status == NOT_RECEIPT and current.detection_status == NOT_RECEIPT and previous.non_receipt_type == current.non_receipt_type:
                _merge_page(previous, current, type_catalog, minimum_confidence)
                continue
            if previous.detection_status == NOT_RECEIPT and previous.non_receipt_type in {"orden_de_compra_pedido", "hoja_entrada_servicios_hes"} and current.detection_status == AMBIGUOUS and continuation_marker:
                _merge_page(previous, current, type_catalog, minimum_confidence)
                continue
            if previous.detection_status == DETECTED and current.detection_status in {NOT_RECEIPT, AMBIGUOUS} and continuation_marker:
                _merge_page(previous, current, type_catalog, minimum_confidence)
                continue
            if previous.detection_status == DETECTED and current.detection_status == DETECTED and _same_complete_identity(previous, current):
                _merge_page(previous, current, type_catalog, minimum_confidence)
                continue
        grouped.append(current)

    # Paso 3: controles finales por resultado, y si necesita revisión humana (indeterminado, error de lectura
    # o comprobante con identidad incompleta o con campos por revisar).
    detected = [item for item in grouped if item.comprobante_detectado]
    can_name_single_result = len(grouped) == 1 and len(detected) == 1
    for item in grouped:
        _compute_derived_checks(item, minimum_confidence)
        item.requires_human_review = (
            item.detection_status in {AMBIGUOUS, READ_ERROR}
            or item.comprobante_detectado and (
                not item.identidad_completa or bool(item.secondary_review_fields)
            )
        )
        item.renombrable = bool(
            can_name_single_result
            and item.identidad_completa
            and not item.read_error
            and min(item.confidence.get(name, 0.0) for name in ("emitter_ruc", "document_type", "series", "correlativo"))
            >= minimum_confidence
        )
    return grouped or [
        ExtractionResult(
            detection_status=AMBIGUOUS,
            identity_status=IDENTITY_NOT_APPLICABLE,
            detection_reason="El archivo no contiene páginas legibles.",
            observations=["el archivo no contiene páginas legibles"],
        )
    ]


# DEPRECADO: nadie lo llama. Era un atajo para analizar una sola página de texto.
def parse_fields(text: str, catalog: Mapping[str, str] | None = None) -> ExtractionResult:
    """Atajo antiguo para analizar una sola página de texto; prioriza los comprobantes detectados si hubo segmentación."""
    results = parse_document([(1, text)], catalog)
    return next((item for item in results if item.comprobante_detectado), results[0])
