from __future__ import annotations

import csv
from dataclasses import dataclass, field
import io
import os
from pathlib import Path
import shutil
import subprocess
from typing import Protocol

from .models import LayoutLine


# Se lanza cuando Tesseract no está disponible: no está instalado, le faltan los idiomas pedidos o no arranca.
class OCRUnavailable(RuntimeError):
    pass


# Se lanza cuando Tesseract sí está, pero falló al leer una página (tiempo agotado o error al ejecutarlo).
class OCRExecutionError(RuntimeError):
    pass


# Lo que devuelve una lectura de OCR: el texto, la confianza media, la versión del motor
# y, línea por línea, su confianza y su posición.
@dataclass(frozen=True)
class OCRResult:
    text: str
    mean_confidence: float | None
    engine_version: str
    line_confidences: dict[str, float] = field(default_factory=dict)
    layout_lines: list[LayoutLine] = field(default_factory=list)


# Contrato mínimo de un motor de OCR. Hoy sólo existe la implementación de Tesseract,
# pero así se podría cambiar por otra.
class OCREngine(Protocol):
    name: str
    version: str

    def availability(self) -> tuple[bool, str | None]: ...

    def recognize(self, image_path: str | Path) -> OCRResult: ...


# El motor de OCR: ejecuta el programa Tesseract en un subproceso y lee su salida en formato TSV.
class TesseractOCREngine:
    name = "tesseract"

    # Guarda los ajustes y los contadores de llamadas. No se ejecuta nada hasta que se pide la disponibilidad o una lectura.
    def __init__(
        self,
        executable: str = "auto",
        languages: str = "spa+eng",
        timeout_seconds: int = 90,
    ) -> None:
        self.executable = executable
        self.languages = languages
        self.timeout_seconds = timeout_seconds
        self._resolved: str | None = None
        self.version = "unknown"
        self._availability_cache: tuple[bool, str | None] | None = None
        # Registramos las sondas de versión/idiomas por separado de las invocaciones de reconocimiento por página.
        self.probe_invocations: int = 0
        self.recognition_invocations: int = 0
        self.probe_events: list[dict[str, object]] = []

    # Busca dónde está Tesseract. Orden: la ruta de config.json, la variable de entorno TESSERACT_CMD,
    # el PATH y, por último, las carpetas habituales de Windows (Program Files).
    def _resolve(self) -> str | None:
        if self._resolved:
            return self._resolved
        if self.executable != "auto":
            candidate = Path(self.executable)
            if candidate.is_file():
                self._resolved = str(candidate)
            else:
                self._resolved = shutil.which(self.executable)
            return self._resolved

        env_override = os.environ.get("TESSERACT_CMD")
        if env_override:
            candidate = Path(env_override)
            if candidate.is_file():
                self._resolved = str(candidate)
                return self._resolved
        self._resolved = shutil.which("tesseract")
        if self._resolved:
            return self._resolved

        program_files = [os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)"), r"C:\Program Files"]
        for root in filter(None, program_files):
            candidate = Path(root) / "Tesseract-OCR" / "tesseract.exe"
            if candidate.is_file():
                self._resolved = str(candidate)
                return self._resolved
        return None

    # Comprueba, una sola vez (el resultado queda en memoria), que Tesseract arranca y tiene los idiomas pedidos.
    # Devuelve (está_disponible, motivo_si_no).
    def availability(self) -> tuple[bool, str | None]:
        if self._availability_cache is not None:
            return self._availability_cache
        # Paso 1: encontrar el ejecutable.
        executable = self._resolve()
        if executable is None:
            self._availability_cache = (False, "Tesseract no está instalado o no se encuentra en PATH")
            return self._availability_cache
        try:
            self.probe_invocations += 1
            # Paso 2: preguntarle la versión. Si no responde bien, se da por no disponible.
            version = subprocess.run(
                [executable, "--version"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=min(self.timeout_seconds, 10),
                check=False,
            )
            self.probe_events.append(
                {
                    "probe": "version",
                    "executable": executable,
                    "returncode": version.returncode,
                    "status": "succeeded" if version.returncode == 0 else "failed",
                }
            )
            if version.returncode != 0:
                self._availability_cache = (False, version.stderr.strip() or "Tesseract no pudo iniciar")
                return self._availability_cache
            self.version = (version.stdout or version.stderr).splitlines()[0].strip()
            self.probe_invocations += 1
            # Paso 3: pedirle la lista de idiomas y comprobar que están todos los que hacen falta (por ejemplo spa y eng).
            languages = subprocess.run(
                [executable, "--list-langs"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=min(self.timeout_seconds, 10),
                check=False,
            )
            available_languages = set(languages.stdout.splitlines()[1:])
            missing = [language for language in self.languages.split("+") if language not in available_languages]
            self.probe_events.append(
                {
                    "probe": "languages",
                    "executable": executable,
                    "returncode": languages.returncode,
                    "status": "succeeded" if languages.returncode == 0 and not missing else "failed",
                    "requested": self.languages,
                    "missing": missing,
                }
            )
            if languages.returncode != 0:
                self._availability_cache = (False, languages.stderr.strip() or "No se pudo leer la lista de idiomas de Tesseract")
                return self._availability_cache
            if missing:
                self._availability_cache = (False, f"Faltan datos de idioma Tesseract: {', '.join(missing)}")
                return self._availability_cache
            self._availability_cache = (True, None)
            return self._availability_cache
        # Cualquier fallo al comprobarlo se anota como «no disponible», con su motivo.
        except (OSError, subprocess.TimeoutExpired) as exc:
            self.probe_events.append(
                {
                    "probe": "availability_exception",
                    "executable": executable,
                    "status": "failed",
                    "error": str(exc),
                }
            )
            self._availability_cache = (False, f"No se pudo comprobar Tesseract: {exc}")
            return self._availability_cache

    # Lee una imagen con Tesseract y devuelve texto, confianza y líneas con su posición.
    # El modo de segmentación 3 es el de página completa; el 6 se usa para recuadros.
    def recognize(self, image_path: str | Path, *, page_segmentation_mode: int = 3) -> OCRResult:
        # Paso 1: si Tesseract no está, se avisa con OCRUnavailable (el programa lo anota y sigue con los PDF de texto).
        available, reason = self.availability()
        if not available:
            raise OCRUnavailable(reason or "Tesseract no está disponible")
        executable = self._resolve()
        assert executable is not None
        self.recognition_invocations += 1
        try:
            # Paso 2: ejecutar Tesseract pidiendo la salida en TSV (una tabla con cada palabra, su caja y su confianza).
            completed = subprocess.run(
                [executable, str(image_path), "stdout", "-l", self.languages, "--psm", str(page_segmentation_mode), "tsv"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=self.timeout_seconds,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise OCRExecutionError(f"OCR excedió {self.timeout_seconds} segundos en {image_path}") from exc
        except OSError as exc:
            raise OCRExecutionError(f"No se pudo ejecutar Tesseract en {image_path}: {exc}") from exc
        if completed.returncode != 0:
            err_detail = completed.stderr.strip() or f"Tesseract terminó con código {completed.returncode}"
            raise OCRExecutionError(err_detail)

        try:
            # Tesseract emite TSV crudo sin escapar comillas dobles; usamos QUOTE_NONE
            # para que una comilla aislada en el OCR no agrupe varias filas TSV en una sola palabra.
            rows = csv.DictReader(io.StringIO(completed.stdout), delimiter="\t", quoting=csv.QUOTE_NONE)
            # Paso 3: agrupar las palabras en líneas (misma página, bloque, párrafo y línea de Tesseract)
            # y guardar, por línea, su posición y su confianza.
            line_words: dict[tuple[int, int, int, int], list[str]] = {}
            line_confidences: dict[tuple[int, int, int, int], list[float]] = {}
            line_boxes: dict[tuple[int, int, int, int], list[tuple[float, float, float, float]]] = {}
            confidences: list[float] = []
            for row in rows:
                word = (row.get("text") or "").strip()
                if not word:
                    continue
                key = tuple(int(row.get(name) or 0) for name in ("page_num", "block_num", "par_num", "line_num"))
                line_words.setdefault(key, []).append(word)
                left = float(row.get("left") or 0)
                top = float(row.get("top") or 0)
                width = float(row.get("width") or 0)
                height = float(row.get("height") or 0)
                if width > 0 and height > 0:
                    line_boxes.setdefault(key, []).append((left, top, left + width, top + height))
                confidence = float(row.get("conf") or -1)
                if confidence >= 0:
                    normalized_confidence = confidence / 100.0
                    confidences.append(normalized_confidence)
                    line_confidences.setdefault(key, []).append(normalized_confidence)
        except (ValueError, TypeError, csv.Error) as exc:
            raise OCRExecutionError("Salida TSV de Tesseract no reconocible") from exc

        # Paso 4: armar el texto final, una línea por grupo. La confianza de la línea es la de su palabra más débil,
        # y su posición es el rectángulo que contiene todas sus palabras.
        output_lines: list[str] = []
        output_line_confidences: dict[str, float] = {}
        layout_lines: list[LayoutLine] = []
        for idx, (key, words) in enumerate(sorted(line_words.items()), start=1):
            line = " ".join(words)
            output_lines.append(line)
            scores = line_confidences.get(key, [])
            line_score: float | None = None
            if scores:
                existing = output_line_confidences.get(line)
                line_score = round(min(scores), 4)
                output_line_confidences[line] = line_score if existing is None else min(existing, line_score)
            boxes = line_boxes.get(key, [])
            bbox: tuple[float, float, float, float] | None = None
            if boxes:
                bbox = (
                    min(b[0] for b in boxes),
                    min(b[1] for b in boxes),
                    max(b[2] for b in boxes),
                    max(b[3] for b in boxes),
                )
            layout_lines.append(
                LayoutLine(
                    line_number=idx,
                    text=line,
                    block_number=max(1, key[1]),
                    bbox=bbox,
                    ocr_confidence=line_score,
                    source="tesseract",
                )
            )
        text = "\n".join(output_lines)
        # La confianza media del OCR es el promedio de todas las palabras reconocidas.
        confidence = sum(confidences) / len(confidences) if confidences else None
        return OCRResult(text, confidence, self.version, output_line_confidences, layout_lines)
