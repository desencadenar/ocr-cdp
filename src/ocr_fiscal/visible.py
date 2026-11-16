"""Leyendas tributarias leídas en el documento y sus etiquetas de salida."""

from collections.abc import Iterator, Mapping, Sequence
import re
from typing import Any


# Reglas para encontrar las leyendas tributarias que están escritas en el documento:
# detracción, retención y buen contribuyente.
# Cada regla tiene cuatro partes: la columna del Excel donde se anota, un código interno,
# el texto que se muestra y la expresión regular que la detecta.
# Sólo se registra lo que está impreso. Si la operación de verdad está sujeta o no a detracción
# es otra cuestión que este programa no decide.
_RULES = (
    # Leyenda oficial del tipo «Operación sujeta al SPOT» o «sujeta a detracción».
    ("detraccion_visible", "leyenda_spot", "Leyenda SPOT",
     r"^\s*(?:operaci[oó]n|factura)\s+sujet[ao]\s+(?:a\s+(?:(?:la|el)\s+)?(?:detracci[oó]n|SPOT)\b|al\s+(?:SPOT\b|Sistema\s+de\s+Pago\s+de\s+Obligaciones\s+Tributarias\b))"),
    # Una línea que rotula un importe como detracción («Monto de detracción: S/ 123.00»).
    ("detraccion_visible", "importe_rotulado", "Importe rotulado como detracción",
     r"^\s*(?:monto|importe|total)\s+(?:de\s+(?:la\s+)?)?detracci[oó]n\s*[:=]?\s*(?:S/\.?|PEN|USD|US\$)?\s*\d"),
    # Mención a que el emisor es agente de retención.
    ("retencion_visible", "leyenda_agente", "Mención a agente de retención",
     r"\bagentes?\s+de\s+retenci[oó]n\b"),
    # Mención de la retención del IGV, o de una operación sujeta a retención.
    ("retencion_visible", "mencion_retencion_igv", "Mención de retención del IGV",
     r"^\s*(?:monto\s+(?:de\s+)?|importe\s+(?:de\s+)?)?(?:retenci[oó]n\s+(?:del\s+)?IGV\b|operaci[oó]n\s+sujet[ao]\s+a\s+(?:la\s+)?retenci[oó]n\b)"),
    # Mención a la condición de buen contribuyente.
    ("buen_contribuyente_visible", "mencion_buen_contribuyente", "Mención a buen contribuyente",
     r"\b(?:buen\s+contribuyente|buenos\s+contribuyentes)\b"),
)
# Compilamos las expresiones una sola vez, sin distinguir mayúsculas de minúsculas.
_PATTERNS = [(field, code, re.compile(pattern, re.I)) for field, code, _, pattern in _RULES]
# Código interno -> texto legible que se muestra en el Excel.
_LABELS = {code: label for _, code, label, _ in _RULES}
# Algunas leyendas empiezan con la palabra «Leyenda:» (a veces con símbolos de tabla que trae el OCR).
# La quitamos antes de buscar para que no estorbe.
_LEGEND_PREFIX = re.compile(r"^[\s|\[\]#0-9.()-]*leyenda\s*[:=–-]\s*", re.I)
# Prefijo del tipo «% Detracción: 12 %» que a veces va pegado delante del importe.
# Se retira para que no tape el rótulo del monto.
_DEDUCTION_RATE = re.compile(r"^\s*%\s*detracci[oó]n\s*[:=]\s*\d+(?:[.,]\d+)?\s*%?\s+", re.I)
# Un renglón que empieza con «Operación» o «Factura» puede ser el comienzo de la leyenda SPOT,
# que a veces se parte en varias líneas.
_SPOT_START = re.compile(r"^\s*(?:operaci[oó]n|factura)\b", re.I)


# Recorre las líneas de un bloque y va devolviendo cada leyenda que encuentra,
# junto con los números de línea donde está.
def find_visible_mentions(lines: Sequence[str]) -> Iterator[tuple[str, str, tuple[int, ...]]]:
    """Busca leyendas explícitas. Los montos, pagos y obligaciones se determinan fuera de esta lectura."""
    # Paso 1: línea por línea.
    for index, line in enumerate(lines):
        # Paso 2: probamos cada regla contra la línea.
        for field, code, pattern in _PATTERNS:
            # El texto a revisar es la línea tal cual, salvo en dos reglas que necesitan limpiar un prefijo primero.
            candidate = line
            if code == "leyenda_spot":
                candidate = _LEGEND_PREFIX.sub("", line, count=1)
            elif code == "importe_rotulado":
                candidate = _DEDUCTION_RATE.sub("", line, count=1)
            # Coincidencia en una sola línea: la anotamos y pasamos a la regla siguiente.
            if pattern.search(candidate):
                yield field, code, (index,)
                continue
            # Paso 3: si no coincidió pero parece el inicio de la leyenda SPOT, probamos a unirla
            # con las dos o tres líneas que siguen.
            if code == "leyenda_spot" and (_SPOT_START.match(candidate) or _LEGEND_PREFIX.match(line)):
                # La leyenda oficial puede partirse en líneas dentro del mismo bloque.
                for size in (2, 3):
                    end = min(index + size, len(lines))
                    if end > index + 1 and pattern.search(" ".join([candidate, *lines[index + 1:end]])):
                        yield field, code, tuple(range(index, end))
                        break


# Convierte los códigos internos de un campo en el texto que se muestra en el Excel,
# sin repetidos y separados por « | ».
def visible_labels(evidence: Mapping[str, Any], field: str) -> str:
    labels = [
        _LABELS[code]
        for item in evidence.get(field, ())
        if (code := item.get("normalized_value")) in _LABELS
    ]
    return " | ".join(dict.fromkeys(labels))


_SPOT_RATE_RE = re.compile(
    r"(?:%\s*(?:DETRACCI[ÓO]N|SPOT)|PORCENTAJE(?:\s+DE\s+(?:DETRACCI[ÓO]N|SPOT))?|DETRACCI[ÓO]N|SPOT)\s*(?:\([^)]*\))?\s*[:=\-]?\s*"
    r"(\d{1,2}(?:[.,]\d{1,2})?)\s*%",
    re.IGNORECASE,
)
_SPOT_RATE_PAREN_RE = re.compile(
    r"(?:DETRACCI[ÓO]N|SPOT)\s*\(\s*(\d{1,2}(?:[.,]\d{1,2})?)\s*%\s*\)",
    re.IGNORECASE,
)
_SPOT_CODE_RE = re.compile(
    r"(?:BIEN\s*(?:O|/)\s*SERVICIO(?:\s+SUJETO\s+A\s+DETRACCI[ÓO]N)?|C[ÓO]D(?:IGO|\.)?\s*(?:DE\s+)?(?:BIEN(?:\s+O\s+SERVICIO)?|SPOT|DETRACCI[ÓO]N))\s*[:=\-]?\s*"
    r"([0-9]{3})(?!\d)",
    re.IGNORECASE,
)
_SPOT_ACCOUNT_RE = re.compile(
    r"(?:CTA\.?\s*(?:DE\s+)?DETRACCI[ÓO]N(?:ES)?|CUENTA\s+(?:DE\s+)?DETRACCI[ÓO]N(?:ES)?|"
    r"CUENTA\s+(?:DEL\s+)?BANCO\s+DE\s+LA\s+NACI[ÓO]N|CTA\.?\s*(?:CTE\.?\s*)?B\.?N\.?|BANCO\s+DE\s+LA\s+NACI[ÓO]N)\s*[:.#\-N°º]*\s*"
    r"([0-9][0-9\- ]{6,22}[0-9])",
    re.IGNORECASE,
)


def extract_spot_details(lines: Sequence[str]) -> dict[str, tuple[str, int, str] | None]:
    """Extrae % de detracción, código de bien/servicio SPOT y cuenta del Banco de la Nación observados.

    Devuelve un dict con claves `rate`, `code`, `account`, cada una con `(valor_normalizado, idx_linea, excerpt)` o None.
    Si hay valores contradictorios para un mismo dato dentro del bloque, lo deja en None (no inventa ni elige).
    """
    rates: list[tuple[str, int, str]] = []
    codes: list[tuple[str, int, str]] = []
    accounts: list[tuple[str, int, str]] = []

    for idx, line in enumerate(lines):
        text = line.strip()
        if not text:
            continue
        for pat in (_SPOT_RATE_RE, _SPOT_RATE_PAREN_RE):
            for m in pat.finditer(text):
                val = m.group(1).replace(",", ".")
                rates.append((val, idx, text))
        for m in _SPOT_CODE_RE.finditer(text):
            codes.append((m.group(1), idx, text))
        if re.search(r"\b(?:DETRACCI[ÓO]N|SPOT)\b", text, re.IGNORECASE):
            for m in re.finditer(r"\bC[ÓO]D(?:IGO|\.)?\s*[:=\-]?\s*([0-9]{3})(?!\d)", text, re.IGNORECASE):
                codes.append((m.group(1), idx, text))
        for m in _SPOT_ACCOUNT_RE.finditer(text):
            acct = re.sub(r"\s+", "", m.group(1)).strip("-")
            digits_only = acct.replace("-", "")
            if 8 <= len(digits_only) <= 16:
                accounts.append((acct, idx, text))

    def _pick_unique(items: list[tuple[str, int, str]]) -> tuple[str, int, str] | None:
        distinct = {v for v, _, _ in items}
        if len(distinct) == 1:
            return items[0]
        return None

    return {
        "rate": _pick_unique(rates),
        "code": _pick_unique(codes),
        "account": _pick_unique(accounts),
    }
