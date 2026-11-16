from __future__ import annotations

from decimal import Decimal
import re
from typing import Sequence

from .models import ExtractionResult, LayoutLine


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
# Incluye también los rótulos propios de los Recibos por Honorarios Electrónicos (02) de SUNAT.
_LABELS = (
    ("no_gravado", _BASE_LABEL + r"NO\s+GRAVAD[AO]S?"),
    ("gravado", _BASE_LABEL + r"(?:GRAVAD[AO]S?|AFECT[AO]S?|BASE\s+IMPONIBLE(?:\s+IGV)?)"),
    ("inafecto", _BASE_LABEL + r"INAFECT[AO]S?"),
    ("exonerado", _BASE_LABEL + r"EXONERAD[AO]S?"),
    ("gratuitas", _BASE_LABEL + r"GRATUIT[AO]S?"),
    ("exportaciones", _BASE_LABEL + r"EXPORTACI[ÓO]N(?:ES)?"),
    ("subtotal", r"SUB\s*TOTAL(?:\s+(?:BRUTO|NETO|VENTAS?))?|TOTAL\s+POR\s+HONORARIOS"),
    ("valor_venta", r"(?:TOTAL\s+)?VALOR\s+(?:DE\s+)?VENTA"),
    ("descuentos", r"(?:TOTAL\s+)?(?:DESCUENTOS?(?:\s+(?:GLOBAL(?:ES)?|TOTAL(?:ES)?))?|DSCTOS?\.?(?:\s+TOTAL(?:ES)?)?)"),
    ("otros_tributos", r"(?:TOTAL\s+)?OTROS\s+TRIBUTOS"),
    ("otros_cargos", r"(?:TOTAL\s+)?(?:OTROS\s+CARGOS|CARGOS\s+GLOBALES)"),
    ("recargo_consumo", r"(?:(?:RECARGO|REC\.?)\s+(?:(?:AL|DE|POR)\s+)?(?:CONSUMO|SERVICIO)|CARGO\s+POR\s+SERVICIO|SERVICIO\s+\(?10\s*%\)?|R\.?C\.?\s*\(?10\s*%\)?)"),
    ("recargo_fise", r"RECARGO\s+FISE(?:\s*\(\s*OP\.?\s*INAFECTA\s*\))?"),
    ("igv", r"(?:TOTAL\s+)?[I1]\.?G\.?V\.?"),
    ("isc", r"(?:TOTAL\s+)?I\.?S\.?C\.?"),
    ("icbper", r"(?:TOTAL\s+)?(?:I\.?C\.?B\.?P\.?E\.?R\.?|IMP(?:UESTO)?\.?\s+(?:A\s+LAS\s+)?BOLSAS\s+PL[ÁA]STICAS|BOLSAS?\s+PL[ÁA]STICAS?)"),
    ("redondeo", r"(?:AJUSTE\s+(?:DE|POR)\s+)?REDONDEO"),
    ("anticipos", r"(?:TOTAL\s+)?ANTICIPOS?(?:\s+TOTAL(?:ES)?)?"),
    ("importe_detraccion", r"(?:(?:MONTO|IMPORTE|TOTAL)\s+(?:DE\s+(?:LA\s+)?)?)?DETRACCI[ÓO]N(?:\s+(?:TOTAL|SPOT))?"),
    (
        "importe_retencion",
        r"(?:(?:MONTO|IMPORTE|TOTAL)\s+(?:DE\s+)?)?RETENCI[ÓO]N(?:\s*\(\s*\d{1,2}(?:[.,]\d{1,2})?\s*%\s*\))?(?:\s+(?:DEL\s+IGV|I\.?R\.?|IMPUESTO\s+A\s+LA\s+RENTA|TOTAL))?",
    ),
    ("importe_percepcion", r"(?:(?:MONTO|IMPORTE|TOTAL)\s+(?:DE\s+(?:LA\s+)?)?)?PERCEPCI[ÓO]N(?:\s+TOTAL)?"),
    (
        "total_a_pagar",
        r"(?:MONTO\s+NETO\s+PENDIENTE(?:\s+DE\s+PAGO)?|TOTAL\s+NETO\s+RECIBIDO|(?:IMPORTE|TOTAL)\s+(?:NETO\s+)?A\s+PAGAR|NETO\s+A\s+PAGAR)",
    ),
    # «total» va al final porque «TOTAL» a secas aparece también dentro de otros rótulos.
    ("total", r"IMPORTE\s+TOTAL|MONTO\s+TOTAL|TOTAL\s+VENTA|TOTAL\s+GENERAL|TOTAL\s+PRESENTE\s+(?:FACTURA|NOTA\s+DE\s+(?:D[EÉ]BITO|CR[EÉ]DITO))|TOTAL"),
)
# Símbolos de moneda que se aceptan. «S7» y «SY» son formas en que el OCR suele leer «S/»,
# y sólo se admiten si hay un espacio detrás.
_CURRENCY = r"(?:S/\.?|S[7Y](?=\s)|US\$|PEN|USD|EUR|€|\$)"
# Símbolo -> código de moneda. El «$» a secas se queda como «$» porque no dice si son soles o dólares.
_CURRENCY_CODES = {"S/": "PEN", "S/.": "PEN", "S7": "PEN", "SY": "PEN", "US$": "USD", "PEN": "PEN", "USD": "USD", "EUR": "EUR", "€": "EUR", "$": "$"}
# Lo que puede haber entre el rótulo y el número: espacios, dos puntos, corchetes, barras o guiones bajos
# (restos de las tablas que trae el OCR).
_SEPARATOR = r"[\s:=\[\]|_]*"
# Un número con miles y decimales, con coma o con punto («1,234.50», «1 234,50», «55»).
_NUMBER = r"(?:\d{1,3}(?:[,.\u00a0 ]\d{3})+(?:[.,]\d{2})?|\d+(?:[.,]\d{1,2})?)"
# La expresión que busca, en una línea, «rótulo + porcentaje opcional + moneda + número (con paréntesis opcionales para negativos)».
# Cada rótulo es un grupo con nombre, así sabemos de qué importe se trata.
_AMOUNT_RE = re.compile(
    r"(?<!\w)(?:" + "|".join(f"(?P<{field}>{label})" for field, label in _LABELS) + r")"
    r"(?!\w)(?:[\s:=(\-]*\d{1,2}(?:(?:[.,]|\s)\d{1,2})?\s*%\s*\)?)?(?:\s+I\.?R\.?)?"
    + _SEPARATOR + r"(?P<sign>[+\-−](?=\s*" + _CURRENCY + r"))?\s*(?P<currency>" + _CURRENCY + r")?"
    + _SEPARATOR + r"(?P<open_p>\(\s*)?(?:-\s+)?(?P<value>[+\-−]?" + _NUMBER + r")(?(open_p)\s*\))(?![\d.,%]|\s*%)",
    re.IGNORECASE,
)
# Expresión para detectar una línea que sólo contiene un rótulo de importe (sin número en la misma cadena)
# y otra para detectar una línea que sólo contiene una cifra (con moneda opcional) en la columna derecha.
_STANDALONE_LABEL_RE = re.compile(
    r"^\s*(?:" + "|".join(f"(?P<{field}>{label})" for field, label in _LABELS) + r")"
    r"(?:[\s:=(\-]*\d{1,2}(?:(?:[.,]|\s)\d{1,2})?\s*%\s*\)?)?(?:\s+I\.?R\.?)?\s*[:=\[\]|_\-]*\s*(?P<label_curr>" + _CURRENCY + r")?\s*$",
    re.IGNORECASE,
)
_STANDALONE_VALUE_RE = re.compile(
    r"^\s*(?P<sign>[+\-−])?\s*(?P<currency>" + _CURRENCY + r")?\s*[:=\[\]|_]*\s*(?P<open_p>\(\s*)?(?:-\s+)?(?P<value>[+\-−]?" + _NUMBER + r")(?(open_p)\s*\))\s*$",
    re.IGNORECASE,
)
# Grupos de campos que se usan al cuadrar el total: las bases de la venta y lo que se suma encima de ellas.
_BASE_FIELDS = ("gravado", "inafecto", "exonerado", "no_gravado", "exportaciones")
_ADDITIONS = ("igv", "isc", "icbper", "otros_tributos", "otros_cargos", "recargo_consumo", "recargo_fise", "redondeo")
# Detracción, retención y percepción: se leen, pero no forman parte del total del comprobante
# (salvo la retención de 4ta categoría en Recibo por Honorarios, que se descuenta del total por honorarios).
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


_NON_TOTAL_PREFIX_RE = re.compile(
    r"(?:\b(?:DESCUENTOS?|ANTICIPOS?|PERCEPCI[ÓO]N|RETENCI[ÓO]N|DETRACCI[ÓO]N|REDONDEO|"
    r"ICBPER|ISC|GRATUIT[AO]S?|EXONERAD[AO]S?|INAFECT[AO]S?|NO\s+GRAVAD[AO]S?|EXPORTACI[ÓO]N(?:ES)?|"
    r"RECARGO|CUOTAS?|HONORARIOS)\b)\s*(?:DE|DEL|AL)?\s*$",
    re.IGNORECASE,
)
_SPOT_IGV_CONTEXT_RE = re.compile(
    r"\b(?:RETENCI[ÓO]N|DETRACCI[ÓO]N)\b(?:\s+(?:DEL|DE|AL|EL))?\s*(?:\d{1,2}\s*%\s*)?$",
    re.IGNORECASE,
)


# Busca importes rotulados en una línea y los va devolviendo como (campo, valor, moneda, dudoso).
def amount_mentions(line: str):
    """Genera (campo, valor, moneda, dudoso) por cada importe rotulado de la línea."""
    # Paso 1: cada coincidencia es un importe. Primero averiguamos de cuál se trata, según el rótulo que coincidió.
    for match in _AMOUNT_RE.finditer(line):
        field = next(field for field, _ in _LABELS if match[field] is not None)
        prefix_text = line[: match.start()]
        # Evitar falsos «IGV» o «TOTAL» cuando forman parte inmediata de una explicación de retención/detracción
        # o cuando un rótulo de descuento/anticipo/etc. precede inmediatamente a la palabra «TOTAL».
        if field == "igv" and _SPOT_IGV_CONTEXT_RE.search(prefix_text):
            continue
        if field == "total" and _NON_TOTAL_PREFIX_RE.search(prefix_text):
            continue
        # Paso 2: la moneda. Si el importe no trae símbolo y es de detracción, retención o percepción,
        # se busca un «Moneda: ...» más adelante en la misma línea.
        currency = _CURRENCY_CODES.get((match["currency"] or "").upper())
        if currency is None and field in SETTLEMENT_FIELDS:
            explicit = _SETTLEMENT_CURRENCY.search(line, match.end())
            if explicit:
                label = explicit[1].upper()
                currency = "PEN" if label.startswith("SOL") else ("USD" if label.startswith(("DOL", "DÓL")) else label)
        # Paso 3: devolvemos el valor con su signo (cambiando el menos tipográfico por el normal; paréntesis = negativo)
        # y la marca que dice si la cifra es dudosa por culpa del símbolo de moneda.
        raw_val = ((match["sign"] or "") + match["value"]).replace("−", "-")
        if match["open_p"] and not raw_val.startswith(("-", "+")):
            raw_val = f"-{raw_val}"
        yield (
            field,
            raw_val,
            currency,
            _cifra_dudosa_por_el_simbolo(line, match),
        )


def pair_columnar_amounts(
    layout_lines: Sequence[LayoutLine] | None,
) -> list[tuple[str, str, str | None, bool, int, int]]:
    """Empareja rótulos de importe y cifras cuando quedaron en bloques/columnas distintas a la misma altura Y.

    En muchas facturas escaneadas (y en PDFs nativos de dos columnas), los rótulos
    («Op. Gravada», «IGV», «Importe Total») quedan en una columna y los importes
    numéricos («100.00», «18.00», «118.00») quedan en otra columna a la derecha.
    Solo emparejamos cuando existe correspondencia 1-a-1 inequívoca en la misma banda
    vertical Y y el valor está a la derecha del rótulo.
    Devuelve una lista de (field, observed_value, currency, dudoso, label_idx, value_idx).
    """
    if not layout_lines:
        return []

    label_candidates: list[tuple[int, str, LayoutLine, str | None]] = []
    value_candidates: list[tuple[int, re.Match[str], LayoutLine]] = []

    for idx, ln in enumerate(layout_lines):
        if not ln.bbox:
            continue
        text = ln.text.strip()
        if not text:
            continue
        # Si la línea ya tiene rótulo + importe juntos, no es un rótulo huérfano.
        if any(amount_mentions(text)):
            continue
        m_label = _STANDALONE_LABEL_RE.match(text)
        if m_label:
            field = next(f for f, _ in _LABELS if m_label[f] is not None)
            lbl_curr = _CURRENCY_CODES.get((m_label["label_curr"] or "").upper())
            label_candidates.append((idx, field, ln, lbl_curr))
            continue
        m_val = _STANDALONE_VALUE_RE.match(text)
        if m_val and re.search(r"[.,]\d{2}$", m_val["value"].strip()):
            value_candidates.append((idx, m_val, ln))

    if not label_candidates or not value_candidates:
        return []

    used_values: set[int] = set()
    paired: list[tuple[str, str, str | None, bool, int, int]] = []

    for lbl_idx, field, lbl_line, lbl_curr in label_candidates:
        assert lbl_line.bbox is not None
        lx0, ly0, lx1, ly1 = lbl_line.bbox
        l_height = max(ly1 - ly0, 4.0)
        l_cy = (ly0 + ly1) / 2.0

        matches_in_band: list[tuple[int, re.Match[str], LayoutLine]] = []
        for val_idx, m_val, val_line in value_candidates:
            if val_idx in used_values or val_line.source != lbl_line.source:
                continue
            assert val_line.bbox is not None
            vx0, vy0, vx1, vy1 = val_line.bbox
            v_height = max(vy1 - vy0, 4.0)
            v_cy = (vy0 + vy1) / 2.0
            tol_y = max(l_height, v_height) * 0.75
            # El valor debe estar a la derecha del rótulo y en la misma banda vertical Y.
            if vx0 >= lx1 - 6.0 and abs(l_cy - v_cy) <= tol_y:
                matches_in_band.append((val_idx, m_val, val_line))

        # Sólo aceptamos si hay exactamente un número en la misma fila a la derecha.
        if len(matches_in_band) == 1:
            val_idx, m_val, val_line = matches_in_band[0]
            used_values.add(val_idx)
            val_curr = _CURRENCY_CODES.get((m_val["currency"] or "").upper()) or lbl_curr
            dudoso = _cifra_dudosa_por_el_simbolo(val_line.text.strip(), m_val)
            observed = ((m_val["sign"] or "") + m_val["value"]).replace("−", "-")
            if m_val["open_p"] and not observed.startswith(("-", "+")):
                observed = f"-{observed}"
            paired.append((field, observed, val_curr, dudoso, lbl_idx, val_idx))

    return paired


def _extract_issue_year(issue_date: str | None) -> int | None:
    if not issue_date:
        return None
    m = re.match(r"^(\d{4})-\d{2}-\d{2}$", issue_date.strip())
    return int(m.group(1)) if m else None


def evaluate_igv_rate(result: ExtractionResult) -> dict:
    """Calcula la tasa efectiva de IGV y la contrasta con tasas legales peruanas vigentes.

    Considera: 18 % general, régimen especial para restaurantes, hoteles y alojamientos turísticos
    (Ley N.° 31556 / Ley N.° 32219 + IPM Ley N.° 32387: 10 % hasta 2025, 10.5 % en 2026 y 15 % en 2027),
    4 % IVAP y 0 % exonerado/inafecto/exportación. Es un control informativo y no inventa tasas.
    """
    doc_type = result.document_type
    if doc_type in {"02", "20", "40"}:
        return {
            "check": "igv_rate_check",
            "status": "no_aplica",
            "consistent": True,
            "effective_rate": None,
        }

    def _dec(val: str | None) -> Decimal | None:
        if val is None or str(val).strip() == "":
            return None
        try:
            return Decimal(str(val))
        except Exception:
            return None

    gravado = _dec(result.gravado)
    subtotal = _dec(result.subtotal)
    valor_venta = _dec(result.valor_venta)
    inafecto = _dec(result.inafecto)
    exonerado = _dec(result.exonerado)
    exportaciones = _dec(result.exportaciones)
    gratuitas = _dec(result.gratuitas)
    igv = _dec(result.igv)
    total = _dec(result.total)

    has_non_taxable = any(v is not None and abs(v) > 0 for v in (inafecto, exonerado, exportaciones))

    # En operaciones 100 % gratuitas el total a pagar es cero; si aparece IGV suele ser solo referencial.
    if gratuitas is not None and abs(gratuitas) > 0 and not (gravado and abs(gravado) > 0):
        if igv is not None and abs(igv) > 0 and (total is None or total == Decimal("0.00")):
            return {
                "check": "igv_rate_check",
                "status": "igv_referencial_gratuita",
                "consistent": True,
                "effective_rate": None,
                "reason": "IGV referencial impreso en transferencia gratuita (no otorga crédito fiscal).",
            }
        if not has_non_taxable and (igv is None or igv == Decimal("0.00")):
            return {
                "check": "igv_rate_check",
                "status": "exonerado_inafecto_sin_igv",
                "consistent": True,
                "effective_rate": 0.0,
            }

    # Elegimos la base gravada sobre la que corresponde calcular la tasa efectiva.
    base_grav: Decimal | None = None
    if gravado is not None and abs(gravado) > 0:
        base_grav = abs(gravado)
    elif not has_non_taxable:
        cand = valor_venta if (valor_venta is not None and abs(valor_venta) > 0) else subtotal
        if cand is not None and abs(cand) > 0 and igv is not None and abs(igv) > 0:
            base_grav = abs(cand)

    if base_grav is None or base_grav == 0:
        if has_non_taxable:
            if igv is not None and abs(igv) > 0:
                return {
                    "check": "igv_rate_check",
                    "status": "igv_con_base_no_gravada",
                    "consistent": False,
                    "effective_rate": None,
                    "reason": "Se leyó IGV mayor a cero sin base gravada legible (solo figuran operaciones exoneradas, inafectas o de exportación).",
                }
            return {
                "check": "igv_rate_check",
                "status": "exonerado_inafecto_sin_igv",
                "consistent": True,
                "effective_rate": 0.0,
            }
        return {
            "check": "igv_rate_check",
            "status": "no_evaluable",
            "consistent": True,
            "effective_rate": None,
        }

    if igv is None:
        return {
            "check": "igv_rate_check",
            "status": "no_evaluable",
            "consistent": True,
            "effective_rate": None,
        }

    igv_abs = abs(igv)
    effective_rate = (igv_abs / base_grav) * Decimal("100")
    effective_rounded = float(round(effective_rate, 2))

    # Tasas conjuntas (IGV + IPM) vigentes:
    # - 18 % régimen general (cambia la composición interna IGV/IPM por Ley 32387, pero siempre suma 18 %).
    # - 4 % IVAP.
    # - Tasa especial para restaurantes, hoteles y alojamientos turísticos (Ley 31556 / Ley 32219 + IPM Ley 32387;
    #   no aplica a cualquier MYPE ni a catering/concesionarios): 10 % hasta 2025 (8+2), 10.5 % en 2026 (8+2.5) y 15 % en 2027 (12+3).
    year = _extract_issue_year(result.issue_date)
    allowed_rates: list[Decimal] = [Decimal("18.0"), Decimal("4.0")]
    if year is None:
        allowed_rates.extend([Decimal("10.0"), Decimal("10.5"), Decimal("15.0")])
    elif year <= 2025:
        allowed_rates.append(Decimal("10.0"))
    elif year == 2026:
        allowed_rates.append(Decimal("10.5"))
        if doc_type in {"07", "08"}:
            allowed_rates.append(Decimal("10.0"))
    elif year == 2027:
        allowed_rates.append(Decimal("15.0"))
        if doc_type in {"07", "08"}:
            allowed_rates.extend([Decimal("10.5"), Decimal("10.0")])
    else:
        if doc_type in {"07", "08"}:
            allowed_rates.extend([Decimal("15.0"), Decimal("10.5"), Decimal("10.0")])

    # También considerar caso en que el ISC forma parte de la base del IGV (IGV = tasa * (base + isc))
    isc_val = abs(_dec(result.isc) or Decimal("0"))
    candidate_bases = [base_grav]
    if isc_val > 0:
        candidate_bases.append(base_grav + isc_val)

    best_match: tuple[Decimal, Decimal, Decimal] | None = None
    for b in candidate_bases:
        tol_amount = max(Decimal("0.03"), b * Decimal("0.002"))
        for rate in allowed_rates:
            expected_igv = b * rate / Decimal("100")
            diff = abs(igv_abs - expected_igv)
            if diff <= tol_amount:
                if best_match is None or diff < best_match[0]:
                    best_match = (diff, rate, b)

    if best_match is not None:
        _, matched_rate, matched_base = best_match
        eff_vs_b = float(round((igv_abs / matched_base) * Decimal("100"), 2))
        return {
            "check": "igv_rate_check",
            "status": "tasa_valida",
            "consistent": True,
            "effective_rate": eff_vs_b,
            "matched_rate": float(matched_rate),
        }

    rates_str = ", ".join(f"{r:g}%" for r in sorted(set(allowed_rates), reverse=True))
    return {
        "check": "igv_rate_check",
        "status": "tasa_igv_atipica",
        "consistent": False,
        "effective_rate": effective_rounded,
        "reason": f"Tasa efectiva de IGV ({effective_rounded:.2f}%) fuera de tasas legales esperadas ({rates_str}). Verifique base gravada, IGV o régimen aplicable.",
    }


# Cuadra los importes leídos: suma las bases y los cargos y lo compara con el total impreso.
# Si falta algo o hay un conflicto, no inventa nada: devuelve el motivo por el que no se pudo cuadrar.
def reconcile_amounts(result: ExtractionResult, *, rounding_tolerance: Decimal | float = Decimal("0.02")) -> dict:
    """Cruza los importes impresos sin convertir un campo ausente en un cero observado."""
    tol = rounding_tolerance if isinstance(rounding_tolerance, Decimal) else Decimal(str(rounding_tolerance))
    # Paso 1: pasamos los importes leídos a Decimal (sin los errores de redondeo de los float) y dejamos fuera los que no se leyeron.
    values = {
        field: Decimal(value)
        for _, field in AMOUNT_COLUMNS
        if (value := getattr(result, field)) is not None
    }
    # Paso 2: ¿qué total vamos a comparar? El «total» impreso. En RHE (02) priorizamos «total_a_pagar» (Neto recibido).
    observed = values.get("total")
    compared_field = "total"
    if result.document_type == "02" and values.get("total_a_pagar") is not None:
        observed = values["total_a_pagar"]
        compared_field = "total_a_pagar"
    elif observed is None and values.get("total_a_pagar") is not None:
        evidence = result.field_evidence.get("total_a_pagar", [])
        if (
            result.document_type == "02"
            or any(
                re.search(r"\b(?:TOTAL|IMPORTE|MONTO)\s+(?:NETO\s+(?:RECIBIDO|PENDIENTE)|A\s+PAGAR)\b|\bNETO\s+A\s+PAGAR\b", item.excerpt, re.I)
                for item in evidence
            )
        ):
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
    if result.document_type == "02":
        relevant.add("importe_retencion")
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

    # Caso especial 1: Recibo por Honorarios (02), donde Total Neto Recibido = Total por Honorarios - Retención IR.
    if result.document_type == "02" and observed is not None:
        gross_rxh = values.get("subtotal") if values.get("subtotal") is not None else (
            values.get("total") if compared_field == "total_a_pagar" else None
        )
        if gross_rxh is not None:
            ret_ir = abs(values.get("importe_retencion", Decimal(0)))
            expected_rxh = gross_rxh - ret_ir
            difference_rxh = expected_rxh - observed
            exact_rxh = difference_rxh == 0
            rounded_rxh = not exact_rxh and abs(difference_rxh) <= tol
            return {
                **check,
                "status": "cuadra" if exact_rxh else ("cuadra_con_redondeo" if rounded_rxh else "no_cuadra"),
                "consistent": exact_rxh or rounded_rxh,
                "expected_total": f"{expected_rxh:.2f}",
                "difference": f"{difference_rxh:.2f}",
            }

    # Caso especial 2: Operación 100% gratuita (Total = 0.00 y solo Op. Gratuitas sin bases cobradas)
    bases = {field: values[field] for field in _BASE_FIELDS if field in values}
    subtotal = values.get("subtotal")
    sale_value = values.get("valor_venta")
    if observed == Decimal("0.00") and not bases and subtotal is None and sale_value is None:
        if values.get("gratuitas") is not None and abs(values["gratuitas"]) > 0:
            igv_val = values.get("igv") or Decimal("0.00")
            note = (
                "Operación gratuita con IGV referencial (no genera crédito fiscal); total 0.00."
                if igv_val > 0
                else "Operación 100% gratuita sin monto a pagar."
            )
            return {
                **check,
                "status": "cuadra",
                "consistent": True,
                "expected_total": "0.00",
                "difference": "0.00",
                "reason": note,
            }
        # Nota de Crédito Tipo 13 (ajuste de fechas/montos de cuotas con importe 0.00)
        if result.document_type == "07" and (
            result.note_type_code == "13"
            or result.installments_count is not None
            or re.search(r"\bCUOTA\b|\bTIPO\s+13\b|\bAJUSTE\s+DE\s+(?:FECHAS|CUOTAS)\b", result.extracted_text or "", re.I)
        ):
            return {
                **check,
                "status": "cuadra",
                "consistent": True,
                "expected_total": "0.00",
                "difference": "0.00",
                "reason": "Nota de crédito por ajuste de cuotas/fechas (importe total 0.00).",
            }

    # Paso 4: buscamos la base de la venta (bases, subtotal y valor de venta).
    if observed is None or (subtotal is None and sale_value is None and not bases):
        return {**check, "reason": "Hace falta el total del documento o una base de venta legible."}

    # Paso 5: casos de coexistencia no gravado/inafecto o recargo FISE/inafecto.
    discount = abs(values.get("descuentos", Decimal(0)))
    anticipos_val = abs(values.get("anticipos", Decimal(0)))
    fise_inside_inafecto = False
    if values.get("no_gravado") and (values.get("inafecto") or values.get("exonerado")):
        sub_sum = values.get("inafecto", Decimal(0)) + values.get("exonerado", Decimal(0))
        if values["no_gravado"] == sub_sum:
            bases.pop("no_gravado", None)
        else:
            return {**check, "reason": "El no gravado puede incluir las operaciones inafectas o exoneradas."}

    # Paso 6: ¿el subtotal ya viene neto de descuento? Si el papel dice «subtotal neto», no se vuelve a descontar.
    evidence = result.field_evidence.get("subtotal", [])
    net_subtotal = any(
        item.field_status != "lectura_pagina_reemplazada"
        and re.search(r"\bSUB\s*TOTAL\s+NETO\b", item.excerpt, re.I)
        for item in evidence
    )
    after_discount = subtotal if subtotal is None or net_subtotal else subtotal - discount
    additions_sum = sum((values[field] for field in _ADDITIONS if field in values), Decimal(0))

    # Paso 7: elegimos la base contra la que se cuadra.
    if bases:
        base = sum(bases.values(), Decimal(0))
        if sale_value is not None and sale_value != base:
            return {**check, "reason": "El valor de venta no coincide con el desglose de bases leído."}
        if subtotal is not None and subtotal not in {base, base + discount, base + anticipos_val, base + discount + anticipos_val}:
            return {**check, "reason": "El subtotal no coincide con el desglose de bases leído."}
        # Si existe descuento sin subtotal bruto ni valor_venta, aceptamos solo si el total concuerda con bases netas o brutas menos descuento.
        if discount and subtotal is None and sale_value is None:
            if abs((base + additions_sum) - observed) > tol and abs((base - discount + additions_sum) - observed) <= tol:
                base = base - discount
        elif discount and subtotal is not None and subtotal != base and after_discount != base and sale_value != base:
            return {**check, "reason": "El documento no permite distinguir si las bases ya incluyen el descuento."}
    elif sale_value is not None:
        if after_discount is not None and sale_value != after_discount and sale_value != subtotal:
            return {**check, "reason": "El valor de venta difiere del subtotal después de descuentos."}
        base = sale_value
    else:
        if discount and subtotal is not None and abs((subtotal + additions_sum) - observed) <= tol:
            base = subtotal
        else:
            base = after_discount

    # Paso 8: total esperado = base + IGV y demás cargos, evaluando casos especiales (FISE incluido, ISC incluido, anticipos, percepción/detracción en Neto a pagar, NC signos mixtos).
    expected = base + additions_sum
    extra_notes: list[str] = []

    # Caso FISE potencialmente incluido en operaciones inafectas
    if values.get("recargo_fise") and values.get("inafecto"):
        without_fise = expected - values["recargo_fise"]
        if abs(expected - observed) > tol and abs(without_fise - observed) <= tol:
            expected = without_fise
            fise_inside_inafecto = True
            extra_notes.append("Recargo FISE incluido en operaciones inafectas.")
        elif abs(expected - observed) > tol:
            return {**check, "reason": "El recargo FISE puede estar incluido en el total de operaciones inafectas."}

    # Caso ISC ya incluido en base gravada (combustibles/bebidas)
    if values.get("isc") and abs(expected - observed) > tol:
        without_isc = expected - values["isc"]
        if abs(without_isc - observed) <= tol:
            expected = without_isc
            extra_notes.append("ISC incluido dentro de la base gravada.")

    # Caso Anticipos deducidos al pie
    if anticipos_val and abs(expected - observed) > tol and abs((expected - anticipos_val) - observed) <= tol:
        expected = expected - anticipos_val
        extra_notes.append("Anticipo deducido del total.")

    # Caso Total a Pagar con percepción sumada o detracción restada
    if compared_field == "total_a_pagar" and abs(expected - observed) > tol:
        perc_val = abs(values.get("importe_percepcion", Decimal(0)))
        det_val = abs(values.get("importe_detraccion", Decimal(0)))
        if perc_val and abs((expected + perc_val) - observed) <= tol:
            expected = expected + perc_val
            extra_notes.append("Neto a pagar incluye percepción.")
        elif det_val and abs((expected - det_val) - observed) <= tol:
            expected = expected - det_val
            extra_notes.append("Neto a pagar deduce detracción.")

    # Caso Notas de Crédito (07) impresas con signos mixtos (p. ej. bases positivas/negativas con IGV positivo y total negativo)
    effective_observed = observed
    if result.document_type == "07":
        all_signed = [*bases.values(), *(values[f] for f in _ADDITIONS if f in values), observed]
        if any(v > 0 for v in all_signed) and any(v < 0 for v in all_signed):
            if observed < 0 and expected > 0:
                effective_observed = abs(observed)
                extra_notes.append("Nota de crédito con signos mixtos (bases positivas y total negativo).")
            elif abs(expected - observed) > tol:
                abs_base = sum((abs(v) for v in bases.values()), Decimal(0))
                if abs_base == 0 and sale_value is not None:
                    abs_base = abs(sale_value)
                abs_add = sum((abs(values[f]) for f in _ADDITIONS if f in values), Decimal(0))
                for cand_exp in (abs_base + abs_add - abs(discount) - abs(anticipos_val), abs_base + abs_add):
                    if abs(cand_exp - abs(observed)) <= tol:
                        expected = cand_exp
                        effective_observed = abs(observed)
                        extra_notes.append("Nota de crédito con signos mixtos en bases, IGV o total.")
                        break

    difference = expected - effective_observed
    exact = difference == 0
    rounded_ok = not exact and abs(difference) <= tol
    status = "cuadra" if exact else ("cuadra_con_redondeo" if rounded_ok else "no_cuadra")
    res_check = {
        **check,
        "status": status,
        "consistent": exact or rounded_ok,
        "expected_total": f"{expected:.2f}",
        "difference": f"{difference:.2f}",
    }
    if extra_notes:
        res_check["reason"] = " ".join(extra_notes)
    elif rounded_ok:
        res_check["reason"] = f"Cuadra dentro de tolerancia de redondeo ({difference:+.2f})."
    return res_check
