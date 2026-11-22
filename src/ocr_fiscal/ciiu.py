"""Giro del emisor (CIIU) leído de un archivo local de la consulta RUC múltiple de SUNAT.

Este módulo no consulta la red. El usuario genera el archivo con su herramienta `consulta_ruc_masivo`
(`consulta_ruc.py`: Excel con la hoja «RUCs_SUNAT», o el TXT del ZIP de SUNAT con separador «|») y el
programa solo lo lee.

Con el CIIU del emisor se decide si una tasa especial de IGV de restaurantes y hoteles procede.

Fuentes (consultadas el 2026-11-18):
- Ley 31556 (2022) y su reglamento, D.S. 237-2022-EF, art. 2: el régimen cubre a la micro y pequeña
  empresa cuya actividad principal es la clase CIIU Rev. 4 5510 (alojamiento para estancias cortas:
  hoteles y alojamientos turísticos) o 5610 (restaurantes y servicio móvil de comidas). El reglamento
  deja fuera el suministro de comidas por encargo (5621, catering) y otras actividades de servicio de
  comidas (5629, concesionarios), aunque la ley los mencionaba.
- Vigencia: desde el 1 de septiembre de 2022.
- Ley 32219 amplía el régimen y Ley 32387 sube el IPM. Tasa conjunta IGV + IPM:
  - 10 % (8 + 2) hasta 2025.
  - 10.5 % (8 + 2.5) en 2026.
  - 15 % (12 + 3) en 2027, último año según la orientación de SUNAT (emprender.sunat.gob.pe).
  - Desde 2028 vuelve el 18 % general.
  En octubre de 2026 se informó de un proyecto para extenderlo, sin aprobar a esa fecha. Si se aprueba,
  hay que actualizar `SPECIAL_RATE_BY_YEAR` en amounts.py.

Lo que el programa no puede verificar: que el emisor sea MYPE y que la actividad represente al menos el
70 % de sus ingresos. El motivo de `tasa_especial_valida` lo dice.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
import re
import unicodedata

from .tax import validate_ruc

# Clases CIIU Rev. 4 del régimen especial (D.S. 237-2022-EF, art. 2).
SPECIAL_REGIME_CIIU = {
    "5510": "actividades de alojamiento para estancias cortas (hoteles y alojamientos turísticos)",
    "5610": "actividades de restaurantes y de servicio móvil de comidas",
}
# Clases que el reglamento excluye expresamente.
SPECIAL_REGIME_EXCLUDED_CIIU = {
    "5621": "suministro de comidas por encargo (catering)",
    "5629": "otras actividades de servicio de comidas (concesionarios)",
}


# Las mismas clases cuando la consulta trae la descripción Rev. 4 en lugar del código.
_REV4_DESCRIPTIONS = {
    "ACTIVIDADES DE ALOJAMIENTO PARA ESTANCIAS CORTAS": "5510",
    "ACTIVIDADES DE RESTAURANTES Y DE SERVICIO MOVIL DE COMIDAS": "5610",
    "SUMINISTRO DE COMIDAS POR ENCARGO": "5621",
    "OTRAS ACTIVIDADES DE SERVICIO DE COMIDAS": "5629",
}
# Clases CIIU Rev. 3 que SUNAT muestra en la consulta RUC y que contienen el giro del régimen junto con otros que
# no entran: no permiten decidir y la tasa queda para verificar. Cualquier otra actividad Rev. 3 no corresponde a
# ninguna de las clases Rev. 4 del régimen.
_REV3_COMPATIBLE = {
    "RESTAURANTES, BARES Y CANTINAS": "restaurantes (5610) junto con bares (5630) y suministro de comidas (5621), que no entran",
    "HOTELES, CAMPAMENTOS Y OTROS": "hoteles y alojamientos (5510) junto con campamentos (5520), que no entran",
}
_REV3_COMPATIBLE_CODE_PREFIXES = {"5510": "HOTELES, CAMPAMENTOS Y OTROS", "5520": "RESTAURANTES, BARES Y CANTINAS"}


def classify_activity(value: str | None) -> tuple[str, str | None]:
    """Clasifica una actividad: ("incluido" | "excluido", código), ("rev3_compatible", motivo), ("otro", None) o ("sin_dato", None)."""
    if not value:
        return "sin_dato", None
    code = _REV4_DESCRIPTIONS.get(value, value)
    if code in SPECIAL_REGIME_CIIU:
        return "incluido", code
    if code in SPECIAL_REGIME_EXCLUDED_CIIU:
        return "excluido", code
    if code == "5520":
        # Restaurantes, bares y cantinas en Rev. 3; campamentos en Rev. 4. Sin saber la revisión no se decide.
        return "rev3_compatible", "restaurantes, bares y cantinas en Rev. 3 o campamentos en Rev. 4"
    if re.fullmatch(r"\d{5}", code) and code[:4] in _REV3_COMPATIBLE_CODE_PREFIXES:
        return "rev3_compatible", _REV3_COMPATIBLE[_REV3_COMPATIBLE_CODE_PREFIXES[code[:4]]]
    if value in _REV3_COMPATIBLE:
        return "rev3_compatible", _REV3_COMPATIBLE[value]
    return "otro", None


@dataclass(frozen=True)
class EmitterActivity:
    """Lo que dice la consulta RUC de un emisor. Los códigos CIIU van sin descripción."""

    ruc: str
    principal: str | None
    secondary: tuple[str, ...]
    estado: str | None
    inicio_actividades: str | None
    source: str
    query_date: str | None

    def label(self) -> str:
        parts = [f"{self.principal or 'sin principal'} (principal)"]
        parts += list(self.secondary)
        origin = f"{self.source}, {self.query_date}" if self.query_date else self.source
        return f"{'; '.join(parts)} [{origin}]"


@dataclass
class ConsultaRucReport:
    path: str
    query_date: str | None
    activities: dict[str, EmitterActivity] = field(default_factory=dict)
    rejected_rows: int = 0


def _norm_header(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^A-Z0-9]", "", text.upper())


_CODE_RE = re.compile(r"(?<!\d)(\d{4,5})(?!\d)")


def _activity(value: object) -> str | None:
    """Código CIIU de la celda («5610 - RESTAURANTES» o «5610»); sin código, la descripción normalizada.

    La consulta RUC real (verificada el 2026-11-18 con una salida de consulta_ruc.py) trae solo la descripción
    Rev. 3 de SUNAT, por ejemplo «RESTAURANTES, BARES Y CANTINAS.», y «-» cuando no hay actividad.
    """
    text = str(value or "").strip()
    if not text or set(text) <= {"-", " "}:
        return None
    m = _CODE_RE.search(text)
    if m:
        return m.group(1)
    return _norm_description(text)


def _norm_description(text: str) -> str:
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch)).upper()
    return re.sub(r"\s+", " ", text).strip(" .")


def _date(value: object) -> str | None:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return None


def _read_rows(path: Path) -> list[list[object]]:
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        import openpyxl

        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb["RUCs_SUNAT"] if "RUCs_SUNAT" in wb.sheetnames else wb.worksheets[0]
            return [list(r) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp1252", errors="replace")
    return [row for row in csv.reader(text.splitlines(), delimiter="|")]


def load_consulta_ruc(path: str | Path) -> ConsultaRucReport:
    """Lee el archivo de la consulta RUC. Lanza ValueError si no trae RUC o CIIU reconocibles."""
    path = Path(path)
    rows = [r for r in _read_rows(path) if any(str(c or "").strip() for c in r)]
    if not rows:
        raise ValueError(f"El archivo de consulta RUC {path.name} está vacío.")
    headers = [_norm_header(h) for h in rows[0]]

    def find(pred) -> list[int]:
        return [i for i, h in enumerate(headers) if pred(h)]

    ruc_cols = find(lambda h: h in {"NUMERORUC", "RUC", "NUMERODERUC", "NRORUC", "NUMRUC"})
    activity_cols = find(lambda h: ("CIIU" in h or "ACTIVIDADECONOMICA" in h or h.startswith("ACTIVIDADPRINCIPAL")
                                    or h.startswith("ACTIVIDADSECUNDARIA")) and "COMERCIOEXTERIOR" not in h)
    if not ruc_cols:
        raise ValueError(f"El archivo de consulta RUC {path.name} no tiene la columna NumeroRuc.")
    if not activity_cols:
        raise ValueError(
            f"El archivo de consulta RUC {path.name} no tiene columnas de actividad económica (CIIU). "
            f"Columnas encontradas: {', '.join(str(h) for h in rows[0])}."
        )
    principal_cols = [i for i in activity_cols if "PRINC" in headers[i]]
    secondary_cols = [i for i in activity_cols if i not in principal_cols]
    estado_cols = find(lambda h: h.startswith("ESTADO"))
    inicio_cols = find(lambda h: "INICIO" in h)

    mtime = datetime.fromtimestamp(path.stat().st_mtime).date().isoformat()
    report = ConsultaRucReport(path=str(path), query_date=mtime)
    for row in rows[1:]:
        def cell(i: int) -> object:
            return row[i] if i < len(row) else None

        number = re.sub(r"\D", "", str(cell(ruc_cols[0]) or ""))
        check = validate_ruc(number)
        if not (check.format_valid and check.checksum_valid):
            report.rejected_rows += 1
            continue
        principal_codes = [c for i in principal_cols if (c := _activity(cell(i)))]
        secondary_codes = [c for i in secondary_cols if (c := _activity(cell(i)))]
        # Sin una columna rotulada «principal» no se sabe cuál es la actividad principal.
        principal = principal_codes[0] if principal_codes else None
        secondary = tuple(dict.fromkeys(c for c in (*principal_codes[1:], *secondary_codes) if c != principal))
        report.activities[number] = EmitterActivity(
            ruc=number,
            principal=principal,
            secondary=secondary,
            estado=(str(cell(estado_cols[0])).strip().upper() or None) if estado_cols and cell(estado_cols[0]) else None,
            inicio_actividades=_date(cell(inicio_cols[0])) if inicio_cols else None,
            source=path.name,
            query_date=mtime,
        )
    return report


def _with(chk: dict, status: str, reason: str, act: EmitterActivity | None) -> dict:
    out = dict(chk)
    out["status"] = status
    out["reason"] = reason
    if status == "tasa_especial_valida":
        out["consistent"], out["needs_review"] = True, False
    elif status == "tasa_igv_atipica":
        out["consistent"], out["needs_review"] = False, True
    else:
        out["consistent"], out["needs_review"] = None, True
    if act is not None:
        out["ciiu_emisor"] = act.label()
    return out


# Decisión del usuario (2026-11-20): con una actividad Rev. 3 mixta deciden las descripciones del comprobante.
# Comida → restaurante (5610); alojamiento → hotel (5510). Solo bebidas de bar, campamento o nada legible no deciden.
_FOOD_WORDS = re.compile(
    r"\b(?:MENU|ALMUERZOS?|DESAYUNOS?|CENAS?|ALIMENTOS?|COMIDAS?|PLATOS?|POLLOS?|BRASA|CEVICHES?|CHULETAS?|LOMOS?|"
    r"ARROZ|CHAUFA|SOPAS?|CALDOS?|ENSALADAS?|POSTRES?|PIZZAS?|HAMBURGUESAS?|SANDWICH(?:ES)?|PARRILLAS?|ANTICUCHOS?|"
    r"TALLARIN(?:ES)?|PAPAS?|CARNES?|PESCADOS?|FILETES?|TRUCHAS?|ENTRADAS?|PIQUEOS?|BUFFET|SALTADO|CAUSAS?|"
    r"CHICHARRON(?:ES)?|TAMALES?|EMPANADAS?|GUISOS?|ESTOFADOS?|ADOBOS?|ROCOTOS?|ALPACA|CUY|PASTAS?|RISOTTO|"
    r"SUSHI|MAKIS?|WRAPS?|TACOS?|POSTRE|HELADOS?|TORTAS?|CONSUMO\s+DE\s+(?:ALIMENTOS|COMIDAS?))\b"
)
_BAR_WORDS = re.compile(r"\b(?:CERVEZAS?|PISCOS?|VINOS?|COCTEL(?:ES)?|COCKTAILS?|TRAGOS?|WHISKY|RON|VODKA|CHILCANOS?|"
                        r"SOUR|GIN|TEQUILA|SHOTS?|BAR)\b")
_LODGING_WORDS = re.compile(r"\b(?:HOSPEDAJES?|ALOJAMIENTOS?|HABITACION(?:ES)?|ESTADIAS?|HOTEL|SUITES?|NOCHES?\s+DE\s+HOTEL)\b")
_CAMP_WORDS = re.compile(r"\b(?:CAMPING|CAMPAMENTOS?|CARPAS?)\b")


def _norm_text(text: str | None) -> str:
    return _norm_description(text or "") if text else ""


def _decide_by_descriptions(principal: str, descriptions: str | None) -> tuple[str | None, str]:
    """Con «RESTAURANTES, BARES Y CANTINAS» u «HOTELES, CAMPAMENTOS Y OTROS», ¿qué dicen las descripciones?

    Devuelve (código Rev. 4 confirmado o None, motivo).
    """
    text = _norm_text(descriptions)
    restaurant_side = principal.startswith("RESTAURANTES") or principal[:4] == "5520"
    if not text:
        return None, "no se leyeron descripciones del comprobante"
    if restaurant_side:
        if _FOOD_WORDS.search(text):
            return "5610", "las descripciones del comprobante muestran expendio de comidas"
        if _BAR_WORDS.search(text):
            return None, "las descripciones del comprobante muestran solo bebidas, propias de un bar (5630, fuera del régimen)"
        return None, "las descripciones del comprobante no muestran expendio de comidas"
    if _CAMP_WORDS.search(text):
        return None, "las descripciones del comprobante muestran campamento (5520, fuera del régimen)"
    if _LODGING_WORDS.search(text):
        return "5510", "las descripciones del comprobante muestran alojamiento"
    return None, "las descripciones del comprobante no muestran alojamiento"


def refine_special_rate_check(
    chk: dict,
    act: EmitterActivity | None,
    *,
    issue_date: str | None,
    document_type: str | None,
    descriptions: str | None = None,
) -> dict:
    """Decide una tasa especial con el CIIU del emisor. Solo toca los controles `tasa_especial_verificar`.

    - CIIU principal incluido y tasa del año → `tasa_especial_valida`.
    - CIIU principal excluido o no incluido → `tasa_igv_atipica`.
    - Sin dato, CIIU en Rev. 3, solo actividad secundaria, estado distinto de ACTIVO, inicio de actividades
      posterior al comprobante o nota con la tasa del año anterior → `tasa_especial_verificar` con el motivo.
    """
    if chk.get("status") != "tasa_especial_verificar":
        return chk
    rate = chk.get("matched_rate")
    rate_txt = f"{rate:g} %" if isinstance(rate, (int, float)) else "especial"
    year = int(issue_date[:4]) if issue_date and re.match(r"^\d{4}-", issue_date) else None
    if year is None:
        return chk  # Sin fecha no se sabe qué tasa regía: se mantiene para verificar.
    if act is None:
        return _with(chk, "tasa_especial_verificar", (
            f"Tasa de {rate_txt} en {year}, propia del régimen de restaurantes, hoteles y alojamientos turísticos. "
            "El RUC emisor no figura en la consulta RUC cargada; verifique su giro."
        ), None)

    principal = act.principal
    if act.estado and act.estado != "ACTIVO":
        return _with(chk, "tasa_especial_verificar", (
            f"Tasa de {rate_txt} en {year}. La consulta RUC muestra el estado {act.estado}: el CIIU actual puede no "
            "ser el de la fecha del comprobante. Verifique el giro."
        ), act)
    if act.inicio_actividades and issue_date and act.inicio_actividades > issue_date:
        return _with(chk, "tasa_especial_verificar", (
            f"Tasa de {rate_txt} en {year}. La fecha de inicio de actividades ({act.inicio_actividades}) es posterior "
            "al comprobante: el CIIU de la consulta puede no ser el de esa fecha."
        ), act)
    kind, detail = classify_activity(principal)
    if kind == "sin_dato":
        return _with(chk, "tasa_especial_verificar", (
            f"Tasa de {rate_txt} en {year}. La consulta RUC no trae la actividad principal del emisor; verifique el giro."
        ), act)
    if kind == "excluido":
        return _with(chk, "tasa_igv_atipica", (
            f"Tasa de {rate_txt} en {year}, pero la actividad principal del emisor ({detail}, "
            f"{SPECIAL_REGIME_EXCLUDED_CIIU[detail]}) está excluida del régimen especial (D.S. 237-2022-EF): "
            "el giro del emisor no es restaurante, hotel ni alojamiento."
        ), act)
    by_descriptions = ""
    if kind == "rev3_compatible":
        code, why = _decide_by_descriptions(principal, descriptions)
        if code is None:
            return _with(chk, "tasa_especial_verificar", (
                f"Tasa de {rate_txt} en {year}. La actividad principal del emisor es «{principal}» (CIIU Rev. 3), que "
                f"abarca {detail}; {why}. Verifique el giro."
            ), act)
        kind, detail = "incluido", code
        by_descriptions = f" La actividad del emisor es «{principal}» (CIIU Rev. 3) y {why}."
    if kind == "otro":
        if any(classify_activity(c)[0] in {"incluido", "rev3_compatible"} for c in act.secondary):
            return _with(chk, "tasa_especial_verificar", (
                f"Tasa de {rate_txt} en {year}. La actividad principal del emisor es «{principal}»; restaurante u "
                "hotel figura solo como actividad secundaria y el régimen exige que sea la principal. Verifique."
            ), act)
        return _with(chk, "tasa_igv_atipica", (
            f"Tasa de {rate_txt} en {year}, pero la actividad principal del emisor es «{principal}»: el giro del "
            "emisor no es restaurante, hotel ni alojamiento."
        ), act)
    principal = detail  # Código Rev. 4 del régimen (5510 o 5610).

    from .amounts import special_rate_for_date

    current = special_rate_for_date(issue_date)
    if current is not None and isinstance(rate, (int, float)) and abs(float(current) - float(rate)) < 1e-9:
        return _with(chk, "tasa_especial_valida", (
            f"Tasa de {rate_txt} en {year}, régimen especial: CIIU principal del emisor {principal} "
            f"({SPECIAL_REGIME_CIIU[principal]}).{by_descriptions} No se verificó la condición de MYPE ni el 70 % de ingresos."
        ), act)
    # La tasa coincide con la del año anterior: solo pasa en notas que modifican un comprobante de ese año.
    return _with(chk, "tasa_especial_verificar", (
        f"Nota con la tasa de {rate_txt}, la del año anterior a {year}. El emisor tiene CIIU {principal}; "
        "confirme que modifica un comprobante emitido ese año."
    ), act)
