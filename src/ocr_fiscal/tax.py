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


# Tipos documentales cuyo correlativo o número de boleto/declaración puede superar los 8 dígitos estándar.
_EXTENDED_CORRELATIVO_TYPES = frozenset({"05", "10", "12", "13", "14", "16", "50", "52", "54"})


def validate_dni(value: str) -> bool:
    """Comprueba que un DNI peruano tenga exactamente 8 dígitos ASCII."""
    return bool(re.fullmatch(r"[0-9]{8}", str(value).strip()))


# Deja el correlativo en 8 dígitos con ceros a la izquierda ("169" pasa a "00000169").
# Para boletos/tickets/recibos/DUA que superan 8 dígitos conserva los dígitos sin truncar;
# en tipos estándar (01, 02, 03, 04, 07, 08, 20, 40) exige entre 1 y 8 dígitos.
def normalize_correlativo(value: str | None, *, document_type: str | None = None) -> NormalizationResult:
    # Sin correlativo no hay nada que normalizar.
    if value is None:
        return NormalizationResult(None, None, False, "correlativo ausente")
    # Guardamos el texto original tal cual para poder mostrarlo, y trabajamos con una copia sin espacios.
    original = str(value)
    candidate = original.strip()
    max_digits = 20 if document_type in _EXTENDED_CORRELATIVO_TYPES else 8
    if not re.fullmatch(rf"[0-9]{{1,{max_digits}}}", candidate):
        reason = f"correlativo debe tener entre 1 y {max_digits} dígitos ASCII"
        if len(candidate) > max_digits and candidate.isascii() and candidate.isdigit():
            reason = f"correlativo supera {max_digits} dígitos; no se trunca"
        return NormalizationResult(original, None, False, reason)
    # Si tiene hasta 8 dígitos rellenamos con ceros; si supera 8 (p. ej. boleto aéreo IATA) se conserva completo.
    normalized = candidate.zfill(8) if len(candidate) <= 8 else candidate
    return NormalizationResult(original, normalized, True)
