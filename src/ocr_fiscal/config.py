"""Carga y validación de ajustes de lectura desde config.json."""

from __future__ import annotations

from dataclasses import dataclass, field
import json
from pathlib import Path
import re

from .models import DEFAULT_DOCUMENT_TYPES


# El config.json que viene junto al proyecto es el que se usa si no se indica otro.
DEFAULT_CONFIG = Path(__file__).resolve().parents[2] / "config.json"


# Todos los ajustes del programa en un solo sitio, con sus valores de fábrica.
# load_config() los rellena a partir de config.json.
@dataclass
class AppConfig:
    # Ruta a Tesseract. Con «auto» se busca solo en el equipo.
    tesseract_path: str = "auto"
    # Idiomas del OCR (spa+eng = español e inglés).
    languages: str = "spa+eng"
    # Segundos máximos que se le dan a Tesseract por cada página.
    timeout_seconds: int = 90
    # Confianza mínima (de 0 a 1) para dar por buena una lectura de OCR. Por debajo, el dato va a revisión.
    min_ocr_confidence: float = 0.70
    # Resolución con la que se convierte cada página de PDF en imagen antes del OCR.
    render_dpi: int = 220
    # Tamaño máximo de imagen admitido. Las que lo superan se dejan en revisión.
    max_pixels: int = 40_000_000
    # Si es True, se vuelve a leer el cuadro de importes recortado.
    read_amount_panel: bool = False
    # Tipos de comprobante que se buscan (código SUNAT -> nombre). Por defecto, los de models.py.
    document_types: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_DOCUMENT_TYPES))
    # RUC del comprador con el que se contrastan los comprobantes. Lo escribe el usuario al empezar.
    recipient_ruc: str | None = None
    # Abrir o no los archivos ZIP.
    inspect_archives: bool = True
    # Abrir también los ZIP que vienen dentro de otros ZIP.
    inspect_nested_archives: bool = True
    # Si es True, genera el expediente ordenado en PDF por proveedor (<RUC> - <Nombre>).
    export_pdf_folders: bool = False
    # Longitud máxima de la parte del nombre del proveedor en subcarpetas (para evitar rutas largas en Windows).
    max_provider_folder_name_len: int = 45
    # Extensiones que se tratan como contenedores.
    archive_extensions: tuple[str, ...] = (".zip",)
    # Topes de seguridad al abrir ZIP, para no caer en archivos malintencionados o "bombas":
    # niveles de anidación, número de miembros, bytes descomprimidos en total,
    # bytes por miembro y razón de compresión máxima.
    max_archive_depth: int = 8
    max_members_total: int = 20_000
    max_uncompressed_bytes_total: int = 5_368_709_120
    max_member_bytes: int = 1_073_741_824
    max_compression_ratio: float = 1000.0
    # Si es True (siempre lo es: ningún ajuste de config.json lo cambia), la imagen de cada página se guarda
    # con su resolución anotada, para que Tesseract no suponga una más baja y pierda líneas pequeñas.
    explicit_render_dpi_metadata: bool = True
    # Base local del padrón RUC (SQLite) para contrastar la razón social. None = sin padrón.
    padron_db_path: Path | None = None
    # Archivo local de la consulta RUC múltiple de SUNAT (salida de consulta_ruc.py) con el CIIU de los emisores.
    # Sirve para decidir la tasa especial de restaurantes y hoteles. None = sin consulta.
    consulta_ruc_path: Path | None = None
    # Antigüedad máxima del padrón, en días. None = sin límite (valor por defecto, a pedido del usuario: el padrón
    # de consulta_ruc_masivo se usa aunque sea antiguo). La fecha de corte se informa en la hoja resumen.
    padron_max_age_days: int | None = None
    # Unifica Notas de Crédito/Débito (07/08) dentro de la misma carpeta del proveedor en 01_cdp_por_proveedor.
    unify_notes_by_provider: bool = True
    # Tolerancia máxima de redondeo (en moneda del documento) para aceptar el cuadre de importes como cuadre con redondeo.
    rounding_tolerance: float = 0.02
    # Conservar .ocr_fiscal/state.sqlite3 al terminar. Guarda fragmentos de los documentos (evidencia, RUC,
    # razones sociales, importes, descripciones); con False se borra después de escribir el Excel.
    keep_run_state: bool = True
    # Guardar la lectura completa de cada documento (texto, líneas con coordenadas y confianza) en
    # .ocr_fiscal/lecturas.sqlite3, para que el comando redetectar repita la detección sin OCR.
    keep_full_reading: bool = True


# Lee el archivo de ajustes, lo valida y devuelve un AppConfig.
# Si algo está mal, lanza un ValueError con un mensaje en español que dice qué corregir.
def load_config(path: str | Path | None = None) -> AppConfig:
    # Paso 1: elegir qué archivo leer (el que se indique o el config.json del proyecto) y parsear el JSON.
    config_path = Path(path) if path is not None else DEFAULT_CONFIG
    try:
        raw = json.loads(config_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Revisa el formato de {config_path.name}, línea {exc.lineno}, columna {exc.colno}.") from exc
    # El JSON tiene que ser un objeto { ... }, no una lista ni un valor suelto.
    if not isinstance(raw, dict):
        raise ValueError("La configuración debe empezar con { y terminar con }.")

    # Paso 2: armar los ajustes con sus valores de fábrica, usando los nombres en español de config.json.
    defaults = AppConfig()
    settings = {
        "ruta_tesseract": defaults.tesseract_path,
        "idiomas_ocr": defaults.languages,
        "segundos_por_pagina": defaults.timeout_seconds,
        "umbral_confianza_ocr": int(defaults.min_ocr_confidence * 100),
        "resolucion_pdf": defaults.render_dpi,
        "maximo_pixeles": defaults.max_pixels,
        "releer_recuadro_importes": defaults.read_amount_panel,
        "revisar_archivos_zip": defaults.inspect_archives,
        "exportar_pdf_ordenado": defaults.export_pdf_folders,
        "padron_ruc": None,
        "padron_dias_maximos": defaults.padron_max_age_days,
        "consulta_ruc": None,
        "unificar_notas_en_proveedor": defaults.unify_notes_by_provider,
        "tolerancia_redondeo": defaults.rounding_tolerance,
        "conservar_estado_corrida": defaults.keep_run_state,
        "guardar_lectura_completa": defaults.keep_full_reading,
    }
    # Paso 3: rechazar nombres de ajuste desconocidos (por ejemplo, un error de tipeo) antes de seguir.
    unknown = sorted(raw.keys() - settings.keys())
    if unknown:
        raise ValueError(f"Ajustes desconocidos en {config_path.name}: {', '.join(unknown)}.")
    # Lo que haya escrito el usuario pisa los valores de fábrica.
    settings.update(raw)

    # Paso 4: validar los ajustes numéricos. Deben ser enteros (no true/false, no decimales) y estar dentro de su rango.
    limits = {
        "segundos_por_pagina": (1, None),
        "umbral_confianza_ocr": (0, 100),
        "resolucion_pdf": (72, 600),
        "maximo_pixeles": (1, None),
    }
    for key, (minimum, maximum) in limits.items():
        value = settings[key]
        if type(value) is not int or value < minimum or (maximum is not None and value > maximum):
            permitted = f"entre {minimum} y {maximum}" if maximum is not None else f"mayor o igual a {minimum}"
            raise ValueError(f"{key} debe ser un número entero {permitted}.")

    # Paso 5: validar el resto de ajustes (texto, verdadero/falso y rutas).
    executable = settings["ruta_tesseract"]
    languages = settings["idiomas_ocr"]
    inspect_zips = settings["revisar_archivos_zip"]
    export_pdfs = settings["exportar_pdf_ordenado"]
    read_amount_panel = settings["releer_recuadro_importes"]
    padron = settings["padron_ruc"]
    padron_days = settings["padron_dias_maximos"]
    unify_notes = settings["unificar_notas_en_proveedor"]
    rounding_tol = settings["tolerancia_redondeo"]
    if not isinstance(executable, str) or not executable.strip():
        raise ValueError('ruta_tesseract debe ser "auto" o una ruta escrita entre comillas.')
    if not isinstance(languages, str) or not re.fullmatch(r"[\w/]+(?:\+[\w/]+)*", languages):
        raise ValueError('idiomas_ocr debe tener un formato como "spa" o "spa+eng".')
    if type(inspect_zips) is not bool:
        raise ValueError("revisar_archivos_zip debe ser true o false.")
    if type(export_pdfs) is not bool:
        raise ValueError("exportar_pdf_ordenado debe ser true o false.")
    if type(read_amount_panel) is not bool:
        raise ValueError("releer_recuadro_importes debe ser true o false.")
    if type(unify_notes) is not bool:
        raise ValueError("unificar_notas_en_proveedor debe ser true o false.")
    if padron is not None and (not isinstance(padron, str) or not padron.strip()):
        raise ValueError("padron_ruc debe ser null o una ruta escrita entre comillas.")
    consulta = settings["consulta_ruc"]
    if consulta is not None and (not isinstance(consulta, str) or not consulta.strip()):
        raise ValueError("consulta_ruc debe ser null o una ruta escrita entre comillas.")
    if padron_days is not None and (type(padron_days) is not int or padron_days < 1):
        raise ValueError("padron_dias_maximos debe ser null o un entero mayor o igual a 1.")
    if type(rounding_tol) not in {int, float} or rounding_tol < 0 or rounding_tol > 0.5:
        raise ValueError("tolerancia_redondeo debe ser un número entre 0 y 0.5.")
    if type(settings["conservar_estado_corrida"]) is not bool:
        raise ValueError("conservar_estado_corrida debe ser true o false.")
    if type(settings["guardar_lectura_completa"]) is not bool:
        raise ValueError("guardar_lectura_completa debe ser true o false.")

    # Una ruta relativa se interpreta desde el archivo de ajustes que la declara.
    directory = config_path.resolve().parent
    if executable != "auto":
        executable = str(directory / Path(executable).expanduser())
    # Paso 6: construir el AppConfig final. La confianza pasa de 0-100, como se escribe en el JSON, a 0-1.
    # El interruptor de los ZIP activa a la vez los ZIP y los ZIP anidados.
    return AppConfig(
        tesseract_path=executable,
        languages=languages,
        timeout_seconds=settings["segundos_por_pagina"],
        min_ocr_confidence=settings["umbral_confianza_ocr"] / 100,
        render_dpi=settings["resolucion_pdf"],
        max_pixels=settings["maximo_pixeles"],
        read_amount_panel=read_amount_panel,
        inspect_archives=inspect_zips,
        inspect_nested_archives=inspect_zips,
        export_pdf_folders=export_pdfs,
        padron_db_path=directory / Path(padron).expanduser() if padron is not None else None,
        padron_max_age_days=padron_days,
        consulta_ruc_path=directory / Path(consulta).expanduser() if consulta is not None else None,
        unify_notes_by_provider=unify_notes,
        rounding_tolerance=float(rounding_tol),
        keep_run_state=settings["conservar_estado_corrida"],
        keep_full_reading=settings["guardar_lectura_completa"],
    )
