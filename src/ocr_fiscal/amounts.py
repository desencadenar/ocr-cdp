from __future__ import annotations

from decimal import Decimal
import re

from .models import ExtractionResult


# Los importes que el programa busca en un comprobante.
# Cada par es: cómo se llama la columna en el Excel y el nombre interno del campo.
AMOUNT_COLUMNS = (
    ("Subtotal leído", "subtotal"),
    ("Valor de venta leído", "valor_venta"),
    ("Base gravada", "gravado"),
    ("Base inafecta", "inafecto"),
    ("Base exonerada", "exonerado"),
    ("No gravado sin desglose", "no_gravado"),
    ("Operaciones gratuitas", "gratuitas"),
    ("Exportaciones", "exportaciones"),
    ("Descuentos", "descuentos"),
    ("IGV", "igv"),
    ("ISC", "isc"),
    ("ICBPER", "icbper"),
    ("Otros tributos", "otros_tributos"),
    ("Otros cargos", "otros_cargos"),
    ("Recargo al consumo", "recargo_consumo"),
    ("Recargo FISE", "recargo_fise"),
    ("Redondeo", "redondeo"),
    ("Anticipos", "anticipos"),
    ("Monto de detracción leído", "importe_detraccion"),
    ("Monto de retención leído", "importe_retencion"),
    ("Monto de percepción leído", "importe_percepcion"),
    ("Total leído", "total"),
    ("Neto a pagar leído", "total_a_pagar"),
)
# Sólo los nombres internos, en el mismo orden.
AMOUNT_FIELDS = tuple(field for _, field in AMOUNT_COLUMNS)

# Prefijo opcional que suele ir delante de los rótulos de las bases: «Total», «Valor de venta», «Operaciones» u «Op.».
_BASE_LABEL = r"(?:(?:TOTAL\s+)?(?:VALOR\s+DE\s+VENTA\s*[-–:]?\s*)?(?:OPERACIONES|OP\.?)\s*|TOTAL\s+)?"
# Cómo suele estar escrito cada importe en el papel: una expresión por campo.
# El orden importa: si un rótulo contiene a otro («no gravado» y «gravado»), el más específico va primero.
_LABELS = (
    ("no_gravado", _BASE_LABEL + r"NO\s+GRAVAD[AO]S?"),
    ("gravado", _BASE_LABEL + r"(?:GRAVAD[AO]S?|AFECT[AO]S?)"),
    ("inafecto", _BASE_LABEL + r"INAFECT[AO]S?"),
    ("exonerado", _BASE_LABEL + r"EXONERAD[AO]S?"),
    ("gratuitas", _BASE_LABEL + r"GRATUIT[AO]S?"),
    ("exportaciones", _BASE_LABEL + r"EXPORTACI[ÓO]N(?:ES)?"),
    ("subtotal", r"SUB\s*TOTAL(?:\s+(?:BRUTO|NETO|VENTAS?))?"),
    ("valor_venta", r"(?:TOTAL\s+)?VALOR\s+(?:DE\s+)?VENTA"),
    ("descuentos", r"(?:TOTAL\s+)?(?:DESCUENTOS?(?:\s+(?:GLOBAL(?:ES)?|TOTAL(?:ES)?))?|DSCTOS?\.?)"),
    ("otros_tributos", r"(?:TOTAL\s+)?OTROS\s+TRIBUTOS"),
    ("otros_cargos", r"(?:TOTAL\s+)?(?:OTROS\s+CARGOS|CARGOS\s+GLOBALES)"),
    ("recargo_consumo", r"(?:RECARGO|REC\.?)\s+(?:(?:AL|DE)\s+)?CONSUMO"),
    ("recargo_fise", r"RECARGO\s+FISE(?:\s*\(\s*OP\.?\s*INAFECTA\s*\))?"),
    ("igv", r"(?:TOTAL\s+)?[I1]\.?G\.?V\.?"),
    ("isc", r"(?:TOTAL\s+)?I\.?S\.?C\.?"),
    ("icbper", r"(?:TOTAL\s+)?I\.?C\.?B\.?P\.?E\.?R\.?"),
    ("redondeo", r"(?:AJUSTE\s+(?:DE|POR)\s+)?REDONDEO"),
    ("anticipos", r"(?:TOTAL\s+)?ANTICIPOS?"),
    ("importe_detraccion", r"(?:MONTO|IMPORTE|TOTAL)\s+(?:DE\s+)?DETRACCI[ÓO]N"),
    ("importe_retencion", r"(?:(?:MONTO|IMPORTE|TOTAL)\s+(?:DE\s+)?)?RETENCI[ÓO]N(?:\s+DEL\s+IGV)?"),
    ("importe_percepcion", r"(?:MONTO|IMPORTE|TOTAL)\s+(?:DE\s+)?PERCEPCI[ÓO]N"),
    ("total_a_pagar", r"(?:MONTO\s+NETO\s+PENDIENTE(?:\s+DE\s+PAGO)?|(?:IMPORTE|TOTAL)\s+(?:NETO\s+)?A\s+PAGAR|NETO\s+A\s+PAGAR)"),
    # «total» va al final porque «TOTAL» a secas aparece también dentro de otros rótulos.
    ("total", r"IMPORTE\s+TOTAL|MONTO\s+TOTAL|TOTAL\s+GENERAL|TOTAL\s+PRESENTE\s+(?:FACTURA|NOTA\s+DE\s+(?:D[EÉ]BITO|CR[EÉ]DITO))|TOTAL"),
)
# Símbolos de moneda que se aceptan. «S7» y «SY» son formas en que el OCR suele leer «S/»,
# y sólo se admiten si hay un espacio detrás.
_CURRENCY = r"(?:S/\.?|S[7Y](?=\s)|US\$|PEN|USD|EUR|€|\$)"
# Símbolo -> código de moneda. El «$» a secas se queda como «$» porque no dice si son soles o dólares.
_CURRENCY_CODES = {"S/": "PEN", "S/.": "PEN", "US$": "USD", "PEN": "PEN", "USD": "USD", "EUR": "EUR", "€": "EUR", "$": "$"}
# Lo que puede haber entre el rótulo y el número: espacios, dos puntos, corchetes, barras o guiones bajos
# (restos de las tablas que trae el OCR).
_SEPARATOR = r"[\s:=\[\]|_]*"
# Un número con miles y decimales, con coma o con punto («1,234.50», «1 234,50», «55»).
_NUMBER = r"(?:\d{1,3}(?:[,.\u00a0 ]\d{3})+(?:[.,]\d{2})?|\d+(?:[.,]\d{1,2})?)"
# La expresión que busca, en una línea, «rótulo + porcentaje opcional + moneda + número».
# Cada rótulo es un grupo con nombre, así sabemos de qué importe se trata.
_AMOUNT_RE = re.compile(
    r"(?<!\w)(?:" + "|".join(f"(?P<{field}>{label})" for field, label in _LABELS) + r")"
    r"(?!\w)(?:\s*\(?\s*\d{1,2}(?:(?:[.,]|\s)\d{1,2})?\s*%\s*\)?)?"
    + _SEPARATOR + r"(?P<sign>[+\-−](?=\s*" + _CURRENCY + r"))?\s*(?P<currency>" + _CURRENCY + r")?"
    + _SEPARATOR + r"(?:-\s+)?(?P<value>[+\-−]?" + _NUMBER + r")(?![\d.,%]|\s*%)",
    re.IGNORECASE,
)
# Grupos de campos que se usan al cuadrar el total: las bases de la venta y lo que se suma encima de ellas.
_BASE_FIELDS = ("gravado", "inafecto", "exonerado", "no_gravado", "exportaciones")
_ADDITIONS = ("igv", "isc", "icbper", "otros_tributos", "otros_cargos", "recargo_consumo", "recargo_fise", "redondeo")
# Detracción, retención y percepción: se leen, pero no forman parte del total del comprobante.
SETTLEMENT_FIELDS = ("importe_detraccion", "importe_retencion", "importe_percepcion")
# En estos tres importes la moneda a veces se declara aparte («Moneda: USD»).
_SETTLEMENT_CURRENCY = re.compile(
    r"\bMONEDA\s*[:=]\s*(PEN|USD|EUR|SOL(?:ES)?|D[ÓO]LAR(?:ES)?)\b", re.IGNORECASE
)


# Detecta si, justo después de un número, viene otra cifra separada por un espacio.
_SIGUE_OTRA_CIFRA = re.compile(r"^\s+\d")


def _cifra_dudosa_por_el_simbolo(line: str, match: re.Match[str]) -> bool:
    """¿La cifra arrastra un dígito que en realidad era el símbolo de moneda?

    Tesseract lee a menudo el símbolo del sol, «S/», como «$7», «S7», «[7» o «57». Cuando
    queda separado por un espacio («S/ 55.00» → «S7 55.00») la cifra se lee bien. Hay dos
    casos en los que no:

    - El símbolo queda pegado al número: «S/ 55.00» → «$755.00», y 55 se vuelve 755.
    - Del símbolo sólo sobrevive el 7 y el importe queda detrás: «S/ 0.00» → «[7 0.00»,
      y se lee 7 en lugar de 0.

    En ambos no se puede saber desde el texto qué parte es el importe, así que no damos la
    cifra por leída y la dejamos para revisión. Un «$» separado de la cifra, un «US$» o un
    «S/» bien leído no tienen esta ambigüedad.
    """
    if (match["currency"] or "") == "$" and match.end("currency") == match.start("value"):
        return True
    return match["value"] == "7" and bool(_SIGUE_OTRA_CIFRA.match(line[match.end("value"):]))


# Busca importes rotulados en una línea y los va devolviendo como (campo, valor, moneda, dudoso).
def amount_mentions(line: str):
    """Genera (campo, valor, moneda, dudoso) por cada importe rotulado de la línea."""
    # Paso 1: cada coincidencia es un importe. Primero averiguamos de cuál se trata, según el rótulo que coincidió.
    for match in _AMOUNT_RE.finditer(line):
        field = next(field for field, _ in _LABELS if match[field] is not None)
        # Algunas impresiones colocan "IGV:" en la explicación de una retención o detracción.
        if field in {"igv", "total"} and re.search(r"\b(?:RETENCI[ÓO]N|DETRACCI[ÓO]N)\b", line[:match.start()], re.I):
            continue
        # Paso 2: la moneda. Si el importe no trae símbolo y es de detracción, retención o percepción,
        # se busca un «Moneda: ...» más adelante en la misma línea.
        currency = _CURRENCY_CODES.get((match["currency"] or "").upper())
        if currency is None and field in SETTLEMENT_FIELDS:
            explicit = _SETTLEMENT_CURRENCY.search(line, match.end())
            if explicit:
                label = explicit[1].upper()
                currency = "PEN" if label.startswith("SOL") else ("USD" if label.startswith(("DOL", "DÓL")) else label)
        # Paso 3: devolvemos el valor con su signo (cambiando el menos tipográfico por el normal)
        # y la marca que dice si la cifra es dudosa por culpa del símbolo de moneda.
        yield (
            field,
            ((match["sign"] or "") + match["value"]).replace("−", "-"),
            currency,
            _cifra_dudosa_por_el_simbolo(line, match),
        )


# Cuadra los importes leídos: suma las bases y los cargos y lo compara con el total impreso.
# Si falta algo o hay un conflicto, no inventa nada: devuelve el motivo por el que no se pudo cuadrar.
def reconcile_amounts(result: ExtractionResult) -> dict:
    """Cruza los importes impresos sin convertir un campo ausente en un cero observado."""
    # Paso 1: pasamos los importes leídos a Decimal (sin los errores de redondeo de los float) y dejamos fuera los que no se leyeron.
    values = {
        field: Decimal(value)
        for _, field in AMOUNT_COLUMNS
        if (value := getattr(result, field)) is not None
    }
    # Paso 2: ¿qué total vamos a comparar? El «total» impreso. Si no hay, el «total a pagar», pero sólo
    # cuando la propia página lo rotula así y no hay anticipos de por medio.
    observed = values.get("total")
    compared_field = "total"
    if observed is None and values.get("total_a_pagar") is not None and not values.get("anticipos"):
        # Un importe «total a pagar» puede ser el único total impreso. Un saldo
        # pendiente, en cambio, puede ser el neto de cuotas, detracciones u otros ajustes.
        evidence = result.field_evidence.get("total_a_pagar", [])
        if any(re.search(r"\b(?:TOTAL|IMPORTE)\s+A\s+PAGAR\b", item.excerpt, re.I) for item in evidence):
            observed = values["total_a_pagar"]
            compared_field = "total_a_pagar"
    # Resultado por omisión: datos insuficientes. Los pasos siguientes lo van corrigiendo.
    check = {
        "check": "arithmetic_total_reconciliation",
        "method": "observed_components",
        "status": "datos_insuficientes",
        "observed_total": f"{observed:.2f}" if observed is not None else None,
        "compared_field": compared_field,
        "consistent": None,
    }
    # Paso 3: comprobaciones previas. Si hay más de un valor candidato para un mismo concepto, no se cuadra.
    relevant = {"subtotal", "valor_venta", "total", compared_field, "descuentos", *_BASE_FIELDS, *_ADDITIONS}
    if any(result.candidates.get(field) for field in relevant):
        return {**check, "status": "en_conflicto", "reason": "Hay importes distintos para un mismo concepto."}
    # Si los componentes están en monedas distintas (o en «$» mezclado con soles), tampoco se puede cuadrar.
    currencies = {
        item.decision_comment.removeprefix("Moneda del importe: ")
        for field in relevant
        for item in result.field_evidence.get(field, [])
        if item.field_status != "lectura_pagina_reemplazada" and item.decision_comment.startswith("Moneda del importe: ")
    }
    if len(currencies - {"$"}) > 1 or ("$" in currencies and ("PEN" in currencies or result.currency == "PEN")):
        return {**check, "reason": "Los componentes del total muestran más de una moneda."}
    # Paso 4: buscamos la base de la venta (bases, subtotal y valor de venta).
    # Sin total, o sin ninguna base, faltan datos.
    bases = {field: values[field] for field in _BASE_FIELDS if field in values}
    subtotal = values.get("subtotal")
    sale_value = values.get("valor_venta")
    if observed is None or (subtotal is None and sale_value is None and not bases):
        return {**check, "reason": "Hace falta el total del documento o una base de venta legible."}

    # Paso 5: hay casos en que el documento es ambiguo (un «no gravado» que podría incluir lo inafecto,
    # el recargo FISE dentro de lo inafecto). Ahí se prefiere no cuadrar antes que suponer.
    discount = abs(values.get("descuentos", Decimal(0)))
    if values.get("no_gravado") and (values.get("inafecto") or values.get("exonerado")):
        return {**check, "reason": "El no gravado puede incluir las operaciones inafectas o exoneradas."}
    if values.get("recargo_fise") and values.get("inafecto"):
        return {**check, "reason": "El recargo FISE puede estar incluido en el total de operaciones inafectas."}
    # Paso 6: ¿el subtotal ya viene neto de descuento? Si el papel dice «subtotal neto», no se vuelve a descontar.
    evidence = result.field_evidence.get("subtotal", [])
    net_subtotal = any(
        item.field_status != "lectura_pagina_reemplazada"
        and re.search(r"\bSUB\s*TOTAL\s+NETO\b", item.excerpt, re.I)
        for item in evidence
    )
    after_discount = subtotal if subtotal is None or net_subtotal else subtotal - discount
    # Paso 7: elegimos la base contra la que se cuadra. Con desglose de bases, éstas tienen que ser coherentes
    # con el valor de venta y con el subtotal; si no lo son, se explica la discrepancia en lugar de forzarla.
    if bases:
        base = sum(bases.values(), Decimal(0))
        if sale_value is not None and sale_value != base:
            return {**check, "reason": "El valor de venta no coincide con el desglose de bases leído."}
        if subtotal is not None and subtotal not in {base, base + discount}:
            return {**check, "reason": "El subtotal no coincide con el desglose de bases leído."}
        if discount and after_discount != base and sale_value != base:
            return {**check, "reason": "El documento no permite distinguir si las bases ya incluyen el descuento."}
        # El desglose ya refleja el descuento cuando coincide con el subtotal menos el descuento.
    elif sale_value is not None:
        if after_discount is not None and sale_value != after_discount:
            return {**check, "reason": "El valor de venta difiere del subtotal después de descuentos."}
        base = sale_value
    else:
        base = after_discount
    # Paso 8: total esperado = base + IGV y demás cargos. Se compara con el total impreso:
    # la diferencia exacta decide entre «cuadra» y «no_cuadra».
    # Que cuadre no demuestra que las cifras leídas sean correctas, sólo que son coherentes entre sí.
    expected = base + sum((values[field] for field in _ADDITIONS if field in values), Decimal(0))
    difference = expected - observed
    return {
        **check,
        "status": "cuadra" if difference == 0 else "no_cuadra",
        "consistent": difference == 0,
        "expected_total": f"{expected:.2f}",
        "difference": f"{difference:.2f}",
    }

