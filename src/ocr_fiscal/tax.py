from __future__ import annotations

import re

from .models import NormalizationResult, RucValidation


# Un RUC son exactamente 11 dígitos: sin letras, sin espacios, sin guiones.
_RUC_RE = re.compile(r"^[0-9]{11}$")
# Pesos con los que se calcula el dígito de control (módulo 11). Se multiplican contra los 10 primeros dígitos.
_RUC_WEIGHTS = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)


# Valida un RUC sólo por su forma y por su dígito de control. No consulta nada:
# que un número pase este filtro no quiere decir que exista, que esté activo ni que SUNAT lo reconozca.
def validate_ruc(value: str) -> RucValidation:
    """Comprueba sólo la forma de 11 dígitos y el dígito de control local (módulo 11).

    No acredita que el número esté registrado, activo ni autorizado por SUNAT.
    """
    # Paso 1: la forma. Si no son 11 dígitos exactos, ni siquiera calculamos el dígito de control.
    candidate = str(value)
    if not _RUC_RE.fullmatch(candidate):
        return RucValidation(
            candidate,
            format_valid=False,
            checksum_valid=None,
            reason="RUC debe contener exactamente 11 dígitos ASCII",
        )

    # Paso 2: el dígito de control. Multiplicamos cada uno de los 10 primeros dígitos por su peso,
    # sumamos todo y miramos cuánto sobra al dividir entre 11.
    weighted_sum = sum(int(digit) * weight for digit, weight in zip(candidate[:10], _RUC_WEIGHTS))
    remainder = weighted_sum % 11
    expected = 11 - remainder
    # Regla del módulo 11: un resultado de 10 se toma como 0 y uno de 11 se toma como 1.
    if expected == 10:
        expected = 0
    elif expected == 11:
        expected = 1
    # El número es correcto si el último dígito del RUC coincide con el que acabamos de calcular.
    valid = int(candidate[-1]) == expected
    return RucValidation(
        candidate,
        format_valid=True,
        checksum_valid=valid,
        reason=None if valid else "dígito de control local no coincide",
    )


# Deja el correlativo en 8 dígitos con ceros a la izquierda ("169" pasa a "00000169").
# Si no es numérico o tiene más de 8 dígitos no lo toca, y explica por qué en el resultado.
def normalize_correlativo(value: str | None) -> NormalizationResult:
    # Sin correlativo no hay nada que normalizar.
    if value is None:
        return NormalizationResult(None, None, False, "correlativo ausente")
    # Guardamos el texto original tal cual para poder mostrarlo, y trabajamos con una copia sin espacios.
    original = str(value)
    candidate = original.strip()
    # Si no son de 1 a 8 dígitos, se rechaza. Un correlativo demasiado largo no se recorta: se avisa.
    if not re.fullmatch(r"[0-9]{1,8}", candidate):
        reason = "correlativo debe tener entre 1 y 8 dígitos ASCII"
        if len(candidate) > 8 and candidate.isascii() and candidate.isdigit():
            reason = "correlativo supera 8 dígitos; no se trunca"
        return NormalizationResult(original, None, False, reason)
    # Caso normal: rellenamos con ceros hasta llegar a 8 dígitos.
    return NormalizationResult(original, candidate.zfill(8), True)

