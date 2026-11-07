"""Prepara el entorno local y abre OCR CDP con el comando elegido."""

from importlib.metadata import PackageNotFoundError, version
import os
from pathlib import Path
import subprocess
import sys
import venv


# Carpeta donde vive este archivo. El entorno virtual y la lista de librerías se buscan aquí.
PROJECT = Path(__file__).resolve().parent
ENVIRONMENT = PROJECT / ".venv"
REQUIREMENTS = PROJECT / "requirements.txt"


# Avisa si la versión de Python no es una de las probadas (3.11 a 3.13).
def _supported_python():
    if not (3, 11) <= sys.version_info[:2] < (3, 14):
        raise ValueError("Para preparar este proyecto usa Python 3.11, 3.12 o 3.13.")


# Compara lo instalado con requirements.txt y devuelve las líneas que faltan o están en otra versión.
def _missing_requirements():
    # El listado usa versiones exactas para que dos instalaciones partan de las mismas librerías.
    missing = []
    for line in REQUIREMENTS.read_text(encoding="utf-8-sig").splitlines():
        requirement = line.split("#", 1)[0].strip()
        if not requirement:
            continue
        name, separator, expected = requirement.partition("==")
        if not separator or not name or not expected:
            raise ValueError(f"Revisa esta línea de requirements.txt: {requirement}.")
        try:
            installed = version(name)
        except PackageNotFoundError:
            installed = None
        if installed != expected:
            missing.append(requirement)
    return missing


# Primero se asegura de estar corriendo dentro de .venv (lo crea si no existe y se vuelve a lanzar con ese Python).
# Después instala lo que falte y abre el programa.
def main():
    # Paso 1: ¿estamos dentro de .venv? Si no, lo preparamos y repetimos el mismo comando con su Python.
    if Path(sys.prefix).resolve() != ENVIRONMENT.resolve():
        python = ENVIRONMENT / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
        if not python.is_file():
            _supported_python()
            print("Preparando la carpeta local .venv…", flush=True)
            venv.EnvBuilder(with_pip=True).create(ENVIRONMENT)
        # Ejecutamos el mismo script con el Python local. Las librerías del equipo quedan aparte.
        return subprocess.run([str(python), str(Path(__file__).resolve()), *sys.argv[1:]]).returncode

    _supported_python()
    # Paso 2: ya dentro de .venv, instalamos las librerías que falten. Sólo la primera vez hace falta internet.
    missing = _missing_requirements()
    if missing:
        print("Instalando las librerías que faltan en .venv…", flush=True)
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "-r", str(REQUIREMENTS)],
            check=True,
        )

    # Paso 3: hacemos visible la carpeta src para poder importar ocr_fiscal y arrancar.
    # Sin argumentos se abre la ventana.
    sys.path.insert(0, str(PROJECT / "src"))
    from ocr_fiscal.cli import main as run_ocr_cdp

    return run_ocr_cdp(sys.argv[1:] or ["ui"])


# Cualquier problema de preparación se traduce en un mensaje corto y un código de salida distinto de cero.
if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except subprocess.CalledProcessError:
        print("No se pudieron instalar las librerías. Revisa el mensaje anterior y vuelve a ejecutar iniciar.py.", file=sys.stderr)
        raise SystemExit(1)
    except (OSError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
    except KeyboardInterrupt:
        raise SystemExit(130)
