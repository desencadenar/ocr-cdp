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
