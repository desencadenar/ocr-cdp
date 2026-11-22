"""Interfaz de línea de comandos: análisis, redetección sin OCR, reexportación, calibración, padrón y ventana."""

from __future__ import annotations

import argparse
from pathlib import Path
import sys
from typing import Sequence

from .config import DEFAULT_CONFIG, load_config
from .exports import OUTPUT_FILENAME, export_results
from .processor import FiscalProcessor
from .storage import StateStore
from .tax import validate_ruc


# Define los comandos que acepta `python iniciar.py`: scan, export, calibrar, plantilla-calibracion, preparar-padron y ui.
def _parser() -> argparse.ArgumentParser:
    # El parser principal; cada comando cuelga de él como subcomando.
    parser = argparse.ArgumentParser(
        prog="python iniciar.py",
        description="Detectar comprobantes de pago peruanos y reunir los resultados en Excel.",
    )
    commands = parser.add_subparsers(dest="command", required=True)

    # scan: lee una carpeta de comprobantes y genera CDP.xlsx.
    scan = commands.add_parser("scan", help=f"Leer documentos y generar {OUTPUT_FILENAME}")
    scan.add_argument("--input", type=Path, required=True)
    scan.add_argument("--output", type=Path, required=True)
    scan.add_argument("--recipient-ruc", required=True, help="RUC receptor que se desea contrastar")
    scan.add_argument("--config", type=Path, help="Archivo JSON de ajustes. Por defecto, config.json.")
    scan.add_argument("--padron-db", type=Path, help="Padrón SQLite local para contrastar la razón social")
    scan.add_argument(
        "--consulta-ruc",
        type=Path,
        help="Archivo de la consulta RUC múltiple SUNAT (salida de consulta_ruc.py) con el CIIU de los emisores",
    )
    scan.add_argument("--no-inspect-archives", action="store_true", help="Omitir el contenido de los ZIP")
    scan.add_argument(
        "--exportar-pdf",
        action="store_true",
        help="Generar expediente en PDF ordenado por proveedor dentro de la carpeta de salida",
    )

    # redetectar: repite la detección con las lecturas guardadas de una corrida, sin Tesseract.
    redetect = commands.add_parser(
        "redetectar", help="Repetir la detección y el Excel con las lecturas guardadas de una corrida, sin OCR"
    )
    redetect.add_argument("--output", type=Path, required=True, help="Carpeta de la corrida de origen")
    redetect.add_argument("--destino", type=Path, help="Carpeta nueva del resultado. Por omisión, <corrida>_redetectar")
    redetect.add_argument("--input", type=Path, help="Carpeta de documentos. Por omisión, la de la corrida de origen")
    redetect.add_argument("--recipient-ruc", help="Por omisión, el de la corrida de origen")
    redetect.add_argument("--config", type=Path, help="Archivo JSON de ajustes. Por defecto, config.json.")
    redetect.add_argument("--padron-db", type=Path, help="Padrón SQLite local")
    redetect.add_argument("--consulta-ruc", type=Path, help="Archivo de la consulta RUC múltiple SUNAT")
    redetect.add_argument("--exportar-pdf", action="store_true", help="Generar el expediente PDF")

    # export: vuelve a armar el Excel a partir del estado guardado de una corrida, sin releer los documentos.
    export = commands.add_parser("export", help="Volver a generar el Excel desde el estado de una corrida")
    export.add_argument("--output", type=Path, required=True)

    # calibrar: compara varios ajustes contra una referencia revisada a mano.
    calibration = commands.add_parser("calibrar", help="Comparar perfiles OCR con una referencia revisada")
    calibration.add_argument("--input", type=Path, required=True)
    calibration.add_argument("--reference", type=Path, required=True)
    calibration.add_argument("--output", type=Path, required=True, help="Ruta de calibracion.xlsx")
    calibration.add_argument("--recipient-ruc", required=True)
    calibration.add_argument("--config", type=Path, action="append", help="Repita para comparar perfiles")

    # plantilla-calibracion: crea el Excel vacío donde se anota esa referencia.
    template = commands.add_parser("plantilla-calibracion", help="Crear el Excel para etiquetar la muestra")
    template.add_argument("--output", type=Path, required=True)

    # preparar-padron: convierte el ZIP o TXT oficial del Padrón Reducido SUNAT a una base SQLite local.
    padron_cmd = commands.add_parser("preparar-padron", help="Convertir el Padrón Reducido RUC (TXT o ZIP) a SQLite")
    padron_cmd.add_argument("--input", type=Path, required=True, help="Archivo TXT o ZIP del Padrón Reducido SUNAT")
    padron_cmd.add_argument("--output", type=Path, required=True, help="Ruta de salida .sqlite3")
    padron_cmd.add_argument("--cutoff-date", help="Fecha de corte YYYY-MM-DD (opcional; usa la del archivo si se omite)")
    padron_cmd.add_argument("--encoding", default="cp1252", help="Codificación del archivo fuente (por defecto cp1252)")

    # ui: abre la ventana. Es lo que ocurre si no se escribe ningún comando.
    ui = commands.add_parser("ui", help="Abrir la interfaz gráfica")
    ui.add_argument("--input", type=Path)
    ui.add_argument("--output", type=Path)
    ui.add_argument("--config", type=Path, help="Archivo JSON de ajustes. Por defecto, config.json.")
    return parser


# Comprueba que el RUC del comprador tenga 11 dígitos y un dígito de control correcto. Si no, se detiene.
def _recipient(value: str) -> str:
    value = value.strip()
    result = validate_ruc(value)
    if not (result.format_valid and result.checksum_valid):
        raise ValueError("El RUC receptor debe tener formato y dígito verificador válidos.")
    return value


# redetectar: arma una corrida nueva leyendo cada documento de lecturas.sqlite3 en lugar de pasar por Tesseract.
# Los ajustes que no se indiquen (entrada, RUC objetivo, padrón, consulta RUC, expediente PDF) son los de origen.
def _redetect(args: argparse.Namespace) -> Path:
    from .reading_cache import CACHE_FILENAME

    origin = args.output.resolve()
    readings = origin / ".ocr_fiscal" / CACHE_FILENAME
    state = origin / ".ocr_fiscal" / "state.sqlite3"
    if not readings.is_file():
        raise ValueError(
            f"La corrida {origin} no tiene lecturas guardadas ({CACHE_FILENAME}). "
            "Córrala de nuevo con guardar_lectura_completa en true."
        )
    run = (StateStore(state).get_latest_run_metadata() or {}) if state.is_file() else {}
    meta = run.get("run_config") or {}
    source = (args.input or Path(meta.get("input_root") or "")).resolve() if (args.input or meta.get("input_root")) else None
    if source is None or not source.is_dir():
        raise ValueError("Indique la carpeta de documentos con --input (la corrida de origen no la registra).")
    destination = args.destino
    if destination is None:
        destination = origin.with_name(f"{origin.name}_redetectar")
        n = 2
        while destination.exists():
            destination = origin.with_name(f"{origin.name}_redetectar_{n}")
            n += 1
    destination = destination.resolve()
    if destination == source or source in destination.parents:
        raise ValueError("Elija un destino fuera de la carpeta de documentos.")
    config = load_config(args.config)
    config.recipient_ruc = _recipient(args.recipient_ruc or run.get("target_recipient_ruc") or meta.get("target_recipient_ruc") or "")
    padron = args.padron_db or (Path(meta["padron_ruc"]) if meta.get("padron_ruc") else None)
    if padron is not None:
        config.padron_db_path = padron
    consulta = args.consulta_ruc or (Path(meta["consulta_ruc"]) if meta.get("consulta_ruc") else None)
    if consulta is not None:
        from .ciiu import load_consulta_ruc

        load_consulta_ruc(consulta)
        config.consulta_ruc_path = consulta
    config.inspect_archives = config.inspect_nested_archives = bool(meta.get("inspect_archives", True))
    config.export_pdf_folders = bool(args.exportar_pdf or meta.get("exportar_pdf_ordenado"))
    processor = FiscalProcessor(destination, config, replay_readings=readings)
    try:
        summary = processor.scan(source)
        workbook = export_results(processor.store, destination, summary=summary)
    finally:
        processor.close_readings()
    if not config.keep_run_state:
        processor.discard_state()
    return workbook


# Punto de entrada: interpreta la línea de comandos y ejecuta el comando elegido.
# Devuelve 0 si todo salió bien y 2 si hubo un error de uso.
def main(argv: Sequence[str] | None = None) -> int:
    # Paso 1: leer lo que escribió el usuario.
    args = _parser().parse_args(argv)
    try:
        # scan: leer los documentos y escribir el Excel.
        if args.command == "scan":
            # Resolvemos las rutas y comprobamos que la entrada existe y que la salida no está dentro de ella.
            source = args.input.resolve()
            output = args.output.resolve()
            if not source.is_dir():
                raise ValueError(f"No existe la carpeta de entrada {source}.")
            if output == source or source in output.parents:
                raise ValueError("Elija una carpeta de salida fuera de la entrada.")
            # Cargamos los ajustes (config.json o el archivo indicado) y fijamos el RUC del comprador.
            config = load_config(args.config)
            config.recipient_ruc = _recipient(args.recipient_ruc)
            # Las opciones de la línea de comandos pisan a lo que diga config.json.
            if args.padron_db is not None:
                config.padron_db_path = args.padron_db
            if args.consulta_ruc is not None:
                if not args.consulta_ruc.is_file():
                    raise ValueError(f"No existe el archivo de consulta RUC {args.consulta_ruc}.")
                config.consulta_ruc_path = args.consulta_ruc
            # Se valida la consulta RUC antes del OCR, para no descubrir un archivo inservible al final.
            if config.consulta_ruc_path is not None:
                from .ciiu import load_consulta_ruc

                load_consulta_ruc(config.consulta_ruc_path)
            if args.no_inspect_archives:
                config.inspect_archives = False
                config.inspect_nested_archives = False
            if args.exportar_pdf:
                config.export_pdf_folders = True
            # Leer todos los documentos (el estado queda guardado en la carpeta de salida) y escribir CDP.xlsx.
            processor = FiscalProcessor(output, config)
            summary = processor.scan(source)
            workbook = export_results(processor.store, output, summary=summary)
            processor.close_readings()
            if not config.keep_run_state:
                processor.discard_state()
        elif args.command == "redetectar":
            workbook = _redetect(args)
        # export: reabrir la base de la corrida y regenerar el Excel.
        elif args.command == "export":
            state = args.output / ".ocr_fiscal" / "state.sqlite3"
            if not state.is_file():
                raise ValueError(f"No existe el estado de corrida {state}.")
            workbook = export_results(StateStore(state), args.output)
        # calibrar: se importa aquí dentro para no cargar ese módulo si no se usa.
        elif args.command == "calibrar":
            from .benchmark import run_calibration

            workbook = run_calibration(
                args.input, args.reference, args.output, _recipient(args.recipient_ruc),
                args.config or [DEFAULT_CONFIG],
            )
        # plantilla-calibracion: escribe el Excel vacío para la referencia.
        elif args.command == "plantilla-calibracion":
            from .benchmark import write_reference_template

            workbook = write_reference_template(args.output)
        # preparar-padron: importa el Padrón Reducido SUNAT a una base SQLite local.
        elif args.command == "preparar-padron":
            from .padron import cargar_padron_staging

            padron_res = cargar_padron_staging(
                args.input,
                args.output,
                cutoff_date=args.cutoff_date,
                encoding=args.encoding,
            )
            print(
                f"{padron_res.db_path} ({padron_res.unique_rucs} RUCs válidos, "
                f"rechazados={padron_res.rejected_rows}, fecha_corte={padron_res.cutoff_date})"
            )
            return 0
        else:
            # ui (comando por omisión): abrir la ventana. Si Python no trae Tkinter, se explica cómo usar la terminal.
            config = load_config(args.config)
            try:
                from .ui import launch_ui
            except ImportError as exc:
                raise ValueError("Python no tiene el soporte de ventanas. Puedes ejecutar iniciar.py scan desde la terminal.") from exc

            return launch_ui(input_root=args.input, output_root=args.output, config=config)
        # Mostramos la ruta del Excel generado.
        print(workbook)
        return 0
    # Los errores esperados (ruta mala, ajuste mal escrito...) se muestran en una línea, sin traza.
    except (OSError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
