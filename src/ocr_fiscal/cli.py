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


# Define los comandos que acepta `python iniciar.py`: scan, export, calibrar, plantilla-calibracion y ui.
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
    scan.add_argument("--no-inspect-archives", action="store_true", help="Omitir el contenido de los ZIP")

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
            if args.no_inspect_archives:
                config.inspect_archives = False
                config.inspect_nested_archives = False
            # Leer todos los documentos (el estado queda guardado en la carpeta de salida) y escribir CDP.xlsx.
            processor = FiscalProcessor(output, config)
            summary = processor.scan(source)
            workbook = export_results(processor.store, output, summary=summary)
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
