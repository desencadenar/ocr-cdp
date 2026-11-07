from __future__ import annotations

from dataclasses import dataclass, field
import io
from pathlib import Path, PurePosixPath, PureWindowsPath
import time
import zipfile

from .config import AppConfig
from .models import ArchiveSummary, SourceRef, StageEvent


# El flujo de detección de CDP procesa únicamente representaciones visuales (PDF e imágenes).
# Los archivos XML se inventarían pero se excluyen por formato sin alimentar la identificación.
# Formatos de PDF o imagen que sí se leen como comprobante.
SUPPORTED_DOCUMENT_FORMATS = {
    "pdf",
    "jpg",
    "jpeg",
    "png",
    "tif",
    "tiff",
    "bmp",
    "webp",
    "heic",
}

# Extensión del archivo -> formato. Sirve de respaldo cuando el contenido no permite reconocer el formato.
EXTENSION_TO_FORMAT = {
    ".pdf": "pdf",
    ".jpg": "jpg",
    ".jpeg": "jpg",
    ".png": "png",
    ".tif": "tiff",
    ".tiff": "tiff",
    ".bmp": "bmp",
    ".webp": "webp",
    ".heic": "heic",
    ".xml": "xml",
    ".zip": "zip",
    ".xlsx": "xlsx",
    ".xls": "xls",
    ".docx": "docx",
    ".doc": "doc",
    ".pptx": "pptx",
    ".csv": "csv",
    ".txt": "txt",
    ".json": "json",
}

# Los Excel, Word y PowerPoint modernos son en realidad ZIP por dentro, pero no son contenedores de comprobantes: no se abren.
OFFICE_OPENXML_EXTENSIONS = {".xlsx", ".xlsm", ".docx", ".pptx"}


# Una unidad de trabajo: un documento que hay que leer.
# Está en el disco (disk_path), o en memoria (content_bytes) si salió de un ZIP, o trae un error si no se pudo abrir.
@dataclass
class WorkUnit:
    source_ref: SourceRef
    disk_path: Path | None = None
    content_bytes: bytes | None = None
    error: str | None = None

    @property
    def display_path(self) -> str:
        return self.source_ref.display_location


# Lo que devuelve el recorrido: el inventario de todo lo encontrado, los documentos por leer,
# el resumen de cada ZIP, los eventos y unos contadores.
@dataclass
class WalkResult:
    inventory: list[SourceRef] = field(default_factory=list)
    work_units: list[WorkUnit] = field(default_factory=list)
    archive_summaries: list[ArchiveSummary] = field(default_factory=list)
    events: list[StageEvent] = field(default_factory=list)
    top_level_file_count: int = 0
    supported_unit_count: int = 0
    ignored_unsupported_count: int = 0
    excluded_xml_count: int = 0


# Reconoce el formato de un archivo mirando sus primeros bytes (su «firma»), no su extensión.
# Si la firma no dice nada, usa la extensión como respaldo.
def detect_format_from_bytes(header: bytes, filename: str = "") -> str:
    suffix = Path(filename).suffix.lower()
    if header.startswith(b"%PDF-"):
        return "pdf"
    if header.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if header.startswith(b"\xff\xd8\xff"):
        return "jpg"
    if header.startswith((b"II*\x00", b"MM\x00*")):
        return "tiff"
    if header.startswith(b"BM") and len(header) >= 14:
        return "bmp"
    if len(header) >= 12 and header[:4] == b"RIFF" and header[8:12] == b"WEBP":
        return "webp"
    stripped = header.lstrip(b"\xef\xbb\xbf \t\r\n")
    if stripped.startswith((b"<?xml", b"<Invoice", b"<CreditNote", b"<DebitNote", b"<ubl:")):
        return "xml"
    if header.startswith((b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")):
        if suffix in OFFICE_OPENXML_EXTENSIONS:
            return suffix[1:]
        return "zip"
    if suffix in EXTENSION_TO_FORMAT:
        return EXTENSION_TO_FORMAT[suffix]
    return suffix[1:] if suffix.startswith(".") and len(suffix) > 1 else "unknown"


# ¿Este ZIP es en realidad un archivo de Office? Se mira la extensión o los nombres internos (xl/, word/, ppt/).
def _is_ooxml_package(zf: zipfile.ZipFile, filename: str) -> str | None:
    suffix = Path(filename).suffix.lower()
    if suffix in OFFICE_OPENXML_EXTENSIONS:
        return suffix[1:]
    names = set(zf.namelist())
    if "[Content_Types].xml" in names:
        if any(name.startswith("xl/") for name in names):
            return "xlsx"
        if any(name.startswith("word/") for name in names):
            return "docx"
        if any(name.startswith("ppt/") for name in names):
            return "pptx"
    return None


# Detecta nombres peligrosos dentro de un ZIP: rutas absolutas, con unidad, con «..» (que se saldrían de la carpeta)
# o enlaces simbólicos. Devuelve (es_inseguro, motivo).
def is_unsafe_member_path(member_name: str, info: zipfile.ZipInfo | None = None) -> tuple[bool, str | None]:
    if not member_name or "\x00" in member_name:
        return True, "nombre de miembro vacío o con byte nulo"
    if member_name.startswith(("/", "\\")):
        return True, f"ruta absoluta dentro del ZIP no permitida: {member_name}"
    win_path = PureWindowsPath(member_name)
    posix_path = PurePosixPath(member_name)
    if win_path.drive or posix_path.is_absolute() or win_path.is_absolute():
        return True, f"ruta con unidad o raíz absoluta dentro del ZIP: {member_name}"
    parts = [part for part in member_name.replace("\\", "/").split("/") if part]
    if ".." in parts:
        return True, f"intento de path traversal (..) detectado en miembro ZIP: {member_name}"
    if info is not None:
        mode = (info.external_attr >> 16) & 0o170000
        if mode == 0o120000:
            return True, f"enlace simbólico (symlink) dentro del ZIP bloqueado: {member_name}"
    return False, None


# El recorredor de la entrada: lista los archivos de la carpeta, abre los ZIP (también los que van dentro de otros ZIP,
# con topes de seguridad) y decide qué se lee y qué se deja fuera, dejando constancia de todo en el inventario.
class ContainerWalker:
    """Recorre de forma segura carpetas locales y ZIP anidados."""

    def __init__(self, config: AppConfig | None = None, output_root: Path | None = None) -> None:
        self.config = config or AppConfig()
        self.output_root = output_root.resolve() if output_root is not None else None
        self._next_source_id = 0
        self._total_members_seen = 0
        self._total_uncompressed_bytes = 0

    # Reparte un número de identificación nuevo a cada archivo o miembro que se encuentra.
    def _alloc_id(self) -> int:
        self._next_source_id += 1
        return self._next_source_id

    # Recorre la carpeta de entrada y devuelve un WalkResult con todo lo encontrado.
    def walk(self, input_root: str | Path, *, recursive: bool = True) -> WalkResult:
        # Paso 1: comprobar que la carpeta de entrada existe.
        root = Path(input_root).resolve()
        if not root.is_dir():
            raise ValueError(f"La carpeta de entrada no existe: {root}")
        result = WalkResult()
        iterator = root.rglob("*") if recursive else root.glob("*")
        candidates: list[Path] = []
        # Paso 2: reunir los archivos candidatos. Se saltan los enlaces simbólicos y todo lo que esté
        # dentro de la carpeta de salida, para no leer los resultados propios.
        for path in iterator:
            try:
                if path.is_symlink() or not path.is_file():
                    continue
                resolved = path.resolve()
                if self.output_root is not None and (
                    resolved == self.output_root or self.output_root in resolved.parents
                ):
                    continue
                candidates.append(path)
            except OSError:
                continue

        # Paso 3: ordenarlos por nombre (sin distinguir mayúsculas) para que el resultado salga siempre en el mismo orden,
        # e inspeccionar cada uno.
        candidates.sort(key=lambda item: str(item).casefold())
        result.top_level_file_count = len(candidates)

        for path in candidates:
            self._inspect_disk_file(path, result)
        return result

    # ¿Este formato se lee como comprobante? Los XML nunca (ver el comentario de abajo).
    def _is_supported_format(self, fmt: str) -> bool:
        # Los XML (incluidas facturas UBL) se excluyen siempre de este flujo para evitar que
        # completen campos o incrementen el conteo de comprobantes visibles.
        if fmt == "xml":
            return False
        return fmt in SUPPORTED_DOCUMENT_FORMATS

    # Mira un archivo suelto de la carpeta y decide qué hacer con él: abrirlo como ZIP, leerlo como documento
    # o dejarlo fuera. Siempre queda anotado en el inventario.
    def _inspect_disk_file(self, path: Path, result: WalkResult) -> None:
        started = time.perf_counter()
        source_path_str = str(path.resolve())
        try:
            # Paso 1: leer el tamaño y los primeros 64 bytes. Si el archivo no se puede abrir, se anota el error
            # y se sigue con el siguiente.
            stat_size = path.stat().st_size
            with path.open("rb") as handle:
                header = handle.read(64)
        except OSError as exc:
            sid = self._alloc_id()
            ref = SourceRef(
                source_id=sid,
                source_path=source_path_str,
                detected_format=path.suffix[1:].lower() or "unknown",
                inventory_status="read_error",
                status_reason=f"No se pudo abrir archivo en disco: {exc}",
                action_taken="error_recorded",
            )
            result.inventory.append(ref)
            result.work_units.append(WorkUnit(source_ref=ref, disk_path=path, error=str(exc)))
            result.events.append(
                StageEvent(
                    stage="discovery",
                    status="failed",
                    message=f"Error leyendo archivo {path.name}: {exc}",
                    duration_ms=(time.perf_counter() - started) * 1000.0,
                    source_id=sid,
                )
            )
            return

        # Paso 2: reconocer el formato por su firma.
        detected_fmt = detect_format_from_bytes(header, path.name)
        suffix = path.suffix.lower()

        # Paso 3a: si es un ZIP (y están activados los ZIP), se abre. Los paquetes de Office se dejan fuera.
        if self.config.inspect_archives and (suffix in self.config.archive_extensions or detected_fmt == "zip"):
            if suffix in OFFICE_OPENXML_EXTENSIONS:
                self._record_unsupported(
                    result,
                    source_path=source_path_str,
                    container_chain=(),
                    member_path=None,
                    nested_depth=0,
                    detected_format=suffix[1:],
                    byte_size=stat_size,
                    compressed_size=None,
                    reason=f"Formato {suffix} (contenedor Office OpenXML) no es fuente directa de comprobante fiscal",
                    started=started,
                )
                return
            try:
                raw_bytes = path.read_bytes()
            except OSError as exc:
                sid = self._alloc_id()
                ref = SourceRef(
                    source_id=sid,
                    source_path=source_path_str,
                    detected_format="zip",
                    inventory_status="archive_error",
                    status_reason=f"No se pudo leer contenedor ZIP: {exc}",
                    action_taken="error_recorded",
                    byte_size=stat_size,
                )
                result.inventory.append(ref)
                result.archive_summaries.append(
                    ArchiveSummary(
                        source_id=sid,
                        source_path=source_path_str,
                        corrupted=True,
                        error_reason=str(exc),
                    )
                )
                result.work_units.append(WorkUnit(source_ref=ref, disk_path=path, error=f"ZIP ilegible: {exc}"))
                return

            self._inspect_zip_bytes(
                raw_bytes=raw_bytes,
                source_path=source_path_str,
                container_chain=[path.name],
                member_path=None,
                nested_depth=0,
                byte_size=stat_size,
                compressed_size=None,
                result=result,
            )
            return

        # Paso 3b: si es un PDF o una imagen, se registra como documento por leer.
        if self._is_supported_format(detected_fmt) or self._is_supported_format(suffix[1:] if suffix else ""):
            effective_fmt = detected_fmt if self._is_supported_format(detected_fmt) else suffix[1:].lower()
            sid = self._alloc_id()
            ref = SourceRef(
                source_id=sid,
                source_path=source_path_str,
                container_chain=(),
                member_path=None,
                nested_depth=0,
                detected_format=effective_fmt,
                inventory_status="supported",
                status_reason="Formato documental soportado en disco",
                action_taken="processed",
                byte_size=stat_size,
            )
            result.inventory.append(ref)
            result.work_units.append(WorkUnit(source_ref=ref, disk_path=path))
            result.supported_unit_count += 1
            result.events.append(
                StageEvent(
                    stage="discovery",
                    status="succeeded",
                    message=f"Archivo soportado ({effective_fmt}): {path.name}",
                    duration_ms=(time.perf_counter() - started) * 1000.0,
                    source_id=sid,
                )
            )
        else:
            # Paso 3c: cualquier otro formato se anota como no admitido. Los XML tienen su propio motivo.
            is_xml = detected_fmt == "xml" or suffix == ".xml"
            self._record_unsupported(
                result,
                source_path=source_path_str,
                container_chain=(),
                member_path=None,
                nested_depth=0,
                detected_format="xml" if is_xml else detected_fmt,
                byte_size=stat_size,
                compressed_size=None,
                reason=(
                    "Archivo XML excluido por formato del flujo visual de comprobantes (solo PDF e imágenes)"
                    if is_xml
                    else f"Formato no admitido para extracción directa de comprobante ({detected_fmt})"
                ),
                started=started,
            )

    # Anota en el inventario un archivo que no se va a leer (formato no admitido o XML excluido)
    # y deja un evento con el motivo.
    def _record_unsupported(
        self,
        result: WalkResult,
        *,
        source_path: str,
        container_chain: tuple[str, ...],
        member_path: str | None,
        nested_depth: int,
        detected_format: str,
        byte_size: int,
        compressed_size: int | None,
        reason: str,
        started: float,
    ) -> SourceRef:
        sid = self._alloc_id()
        is_xml = detected_format == "xml"
        inv_status = "excluded_xml" if is_xml else "unsupported"
        action = "excluded_by_format" if is_xml else "ignored_unsupported"
        ref = SourceRef(
            source_id=sid,
            source_path=source_path,
            container_chain=container_chain,
            member_path=member_path,
            nested_depth=nested_depth,
            detected_format=detected_format,
            inventory_status=inv_status,
            status_reason=reason,
            action_taken=action,
            byte_size=byte_size,
            compressed_size=compressed_size,
        )
        result.inventory.append(ref)
        result.ignored_unsupported_count += 1
        if is_xml:
            result.excluded_xml_count += 1
        result.events.append(
            StageEvent(
                stage="discovery",
                status=action,
                message=reason,
                duration_ms=(time.perf_counter() - started) * 1000.0,
                source_id=sid,
                member_path=member_path,
                container_chain=container_chain,
            )
        )
        return ref

    # Abre un ZIP (ya cargado en memoria) y revisa sus miembros uno por uno.
    # Los que son documentos pasan a leerse, los ZIP anidados se abren de nuevo con esta misma función
    # y el resto se anota en el inventario con su motivo.
    def _inspect_zip_bytes(
        self,
        *,
        raw_bytes: bytes,
        source_path: str,
        container_chain: list[str],
        member_path: str | None,
        nested_depth: int,
        byte_size: int,
        compressed_size: int | None,
        result: WalkResult,
        parent_summaries: list[ArchiveSummary] | None = None,
    ) -> None:
        started = time.perf_counter()
        # Paso 1: preparar el resumen de este ZIP. Los ZIP que lo contienen (sus «padres») también acumulan sus cifras.
        container_sid = self._alloc_id()
        summary = ArchiveSummary(
            source_id=container_sid,
            source_path=source_path,
            container_chain=list(container_chain),
            nested_depth=nested_depth,
            max_nested_depth_observed=nested_depth + 1,
        )
        active_summaries = [*(parent_summaries or []), summary]

        try:
            # Paso 2: intentar abrirlo. Si está corrupto, se anota y se termina con este ZIP.
            zf = zipfile.ZipFile(io.BytesIO(raw_bytes))
        except (zipfile.BadZipFile, OSError, ValueError) as exc:
            summary.corrupted = True
            summary.error_reason = f"{type(exc).__name__}: {exc}"
            ref = SourceRef(
                source_id=container_sid,
                source_path=source_path,
                container_chain=tuple(container_chain),
                member_path=member_path,
                nested_depth=nested_depth,
                detected_format="zip",
                inventory_status="archive_error",
                status_reason=f"Contenedor ZIP corrupto o inválido: {exc}",
                action_taken="error_recorded",
                byte_size=byte_size,
                compressed_size=compressed_size,
            )
            result.inventory.append(ref)
            result.archive_summaries.append(summary)
            result.work_units.append(
                WorkUnit(source_ref=ref, error=f"Contenedor ZIP corrupto ({member_path or Path(source_path).name}): {exc}")
            )
            for anc in parent_summaries or []:
                anc.read_errors += 1
            result.events.append(
                StageEvent(
                    stage="archive_inspection",
                    status="failed",
                    message=f"ZIP corrupto: {exc}",
                    duration_ms=(time.perf_counter() - started) * 1000.0,
                    source_id=container_sid,
                    member_path=member_path,
                    container_chain=tuple(container_chain),
                )
            )
            return

        # Paso 3: ya abierto, comprobar que no sea en realidad un archivo de Office.
        with zf:
            ooxml_kind = _is_ooxml_package(zf, member_path or source_path)
            if ooxml_kind is not None:
                self._record_unsupported(
                    result,
                    source_path=source_path,
                    container_chain=tuple(container_chain[:-1]) if member_path else (),
                    member_path=member_path,
                    nested_depth=nested_depth,
                    detected_format=ooxml_kind,
                    byte_size=byte_size,
                    compressed_size=compressed_size,
                    reason=f"Paquete Office OpenXML ({ooxml_kind}) detectado por estructura interna; no es contenedor ZIP de comprobantes",
                    started=started,
                )
                return

            # Paso 4: registrar el ZIP en el inventario como contenedor.
            container_ref = SourceRef(
                source_id=container_sid,
                source_path=source_path,
                container_chain=tuple(container_chain),
                member_path=member_path,
                nested_depth=nested_depth,
                detected_format="zip",
                inventory_status="container",
                status_reason="Contenedor ZIP abierto e inventariado localmente",
                action_taken="inspected_container",
                byte_size=byte_size,
                compressed_size=compressed_size,
            )
            result.inventory.append(container_ref)
            result.archive_summaries.append(summary)

            # Paso 5: comprobar la profundidad. Si pasa del máximo, todos sus miembros se anotan como no procesados
            # y no se abre ninguno.
            infos = zf.infolist()
            member_depth = nested_depth + 1
            if member_depth > self.config.max_archive_depth:
                reason = f"Profundidad ZIP ({member_depth}) supera max_archive_depth ({self.config.max_archive_depth})"
                for anc in active_summaries:
                    anc.limit_reached = True
                    if reason not in anc.limit_reasons:
                        anc.limit_reasons.append(reason)
                for info in infos:
                    if info.is_dir():
                        continue
                    for anc in active_summaries:
                        anc.pending_members.append(info.filename)
                    sid = self._alloc_id()
                    result.inventory.append(
                        SourceRef(
                            source_id=sid,
                            source_path=source_path,
                            container_chain=tuple(container_chain),
                            member_path=info.filename,
                            nested_depth=member_depth,
                            detected_format=Path(info.filename).suffix[1:].lower() or "unknown",
                            inventory_status="archive_limit_reached",
                            status_reason=reason,
                            action_taken="skipped_limit",
                            byte_size=info.file_size,
                            compressed_size=info.compress_size,
                        )
                    )
                result.events.append(
                    StageEvent(
                        stage="archive_inspection",
                        status="archive_limit_reached",
                        message=reason,
                        duration_ms=(time.perf_counter() - started) * 1000.0,
                        source_id=container_sid,
                        member_path=member_path,
                        container_chain=tuple(container_chain),
                    )
                )
                return

            # Paso 6: revisar los miembros uno a uno. Cada uno puede quedar bloqueado, saltado por un tope o leído.
            for index, info in enumerate(infos):
                if info.is_dir():
                    summary.directory_entries += 1
                    sid = self._alloc_id()
                    result.inventory.append(
                        SourceRef(
                            source_id=sid,
                            source_path=source_path,
                            container_chain=tuple(container_chain),
                            member_path=info.filename,
                            nested_depth=member_depth,
                            detected_format="directory",
                            inventory_status="directory",
                            status_reason="Entrada de carpeta interna dentro del ZIP",
                            action_taken="inspected_directory",
                            byte_size=0,
                            compressed_size=0,
                        )
                    )
                    continue

                summary.direct_members += 1
                for anc in active_summaries:
                    anc.recursive_members += 1
                    anc.max_nested_depth_observed = max(anc.max_nested_depth_observed, member_depth)

                # Tope de seguridad: si ya se vieron demasiados miembros en total, los que faltan se anotan como pendientes.
                if self._total_members_seen >= self.config.max_members_total:
                    reason = f"Se alcanzó el límite total de miembros ({self.config.max_members_total})"
                    for anc in active_summaries:
                        anc.limit_reached = True
                        if reason not in anc.limit_reasons:
                            anc.limit_reasons.append(reason)
                    for remaining in infos[index:]:
                        if remaining.is_dir():
                            continue
                        for anc in active_summaries:
                            anc.pending_members.append(remaining.filename)
                        sid = self._alloc_id()
                        result.inventory.append(
                            SourceRef(
                                source_id=sid,
                                source_path=source_path,
                                container_chain=tuple(container_chain),
                                member_path=remaining.filename,
                                nested_depth=member_depth,
                                detected_format=Path(remaining.filename).suffix[1:].lower() or "unknown",
                                inventory_status="archive_limit_reached",
                                status_reason=reason,
                                action_taken="skipped_limit",
                                byte_size=remaining.file_size,
                                compressed_size=remaining.compress_size,
                            )
                        )
                    break

                self._total_members_seen += 1

                # Seguridad: los nombres con «..» o con ruta absoluta se bloquean sin leerlos.
                unsafe, unsafe_reason = is_unsafe_member_path(info.filename, info)
                if unsafe:
                    sid = self._alloc_id()
                    for anc in active_summaries:
                        anc.traversal_blocked_members += 1
                    ref = SourceRef(
                        source_id=sid,
                        source_path=source_path,
                        container_chain=tuple(container_chain),
                        member_path=info.filename,
                        nested_depth=member_depth,
                        detected_format="unsafe_path",
                        inventory_status="archive_traversal_blocked",
                        status_reason=unsafe_reason or "Ruta insegura en miembro ZIP",
                        action_taken="blocked_security",
                        byte_size=info.file_size,
                        compressed_size=info.compress_size,
                    )
                    result.inventory.append(ref)
                    result.events.append(
                        StageEvent(
                            stage="archive_security",
                            status="blocked",
                            message=unsafe_reason or "Ruta bloqueada",
                            source_id=sid,
                            member_path=info.filename,
                            container_chain=tuple(container_chain),
                        )
                    )
                    continue

                # Un miembro con contraseña no se puede leer: se anota el error y se sigue.
                if info.flag_bits & 0x1:
                    sid = self._alloc_id()
                    for anc in active_summaries:
                        anc.encrypted_members += 1
                    ref = SourceRef(
                        source_id=sid,
                        source_path=source_path,
                        container_chain=tuple(container_chain),
                        member_path=info.filename,
                        nested_depth=member_depth,
                        detected_format=Path(info.filename).suffix[1:].lower() or "unknown",
                        inventory_status="archive_encrypted",
                        status_reason="Miembro ZIP cifrado con contraseña; no se puede desencriptar sin credencial",
                        action_taken="error_recorded",
                        byte_size=info.file_size,
                        compressed_size=info.compress_size,
                    )
                    result.inventory.append(ref)
                    result.work_units.append(
                        WorkUnit(
                            source_ref=ref,
                            error=f"Miembro ZIP cifrado ({info.filename}); requiere contraseña",
                        )
                    )
                    continue

                # Topes de seguridad por tamaño: tamaño del miembro, total descomprimido y razón de compresión
                # (para protegerse de las «bombas» ZIP, que ocupan poco comprimidas y muchísimo al abrirse).
                ratio = (
                    float(info.file_size) / float(max(1, info.compress_size))
                    if info.file_size > 0
                    else 1.0
                )
                limit_violation: str | None = None
                if info.file_size > self.config.max_member_bytes:
                    limit_violation = f"Tamaño descomprimido del miembro ({info.file_size} B) supera max_member_bytes ({self.config.max_member_bytes} B)"
                elif self._total_uncompressed_bytes + info.file_size > self.config.max_uncompressed_bytes_total:
                    limit_violation = f"Tamaño descomprimido acumulado supera max_uncompressed_bytes_total ({self.config.max_uncompressed_bytes_total} B)"
                elif ratio > self.config.max_compression_ratio and info.file_size > 1024 * 1024:
                    limit_violation = f"Razón de compresión ({ratio:.1f}) supera max_compression_ratio ({self.config.max_compression_ratio})"
                elif ratio > self.config.max_compression_ratio:
                    limit_violation = f"Razón de compresión ({ratio:.1f}) supera max_compression_ratio ({self.config.max_compression_ratio})"

                if limit_violation is not None:
                    sid = self._alloc_id()
                    for anc in active_summaries:
                        anc.limit_reached = True
                        if limit_violation not in anc.limit_reasons:
                            anc.limit_reasons.append(limit_violation)
                        anc.pending_members.append(info.filename)
                    ref = SourceRef(
                        source_id=sid,
                        source_path=source_path,
                        container_chain=tuple(container_chain),
                        member_path=info.filename,
                        nested_depth=member_depth,
                        detected_format=Path(info.filename).suffix[1:].lower() or "unknown",
                        inventory_status="archive_limit_reached",
                        status_reason=limit_violation,
                        action_taken="skipped_limit",
                        byte_size=info.file_size,
                        compressed_size=info.compress_size,
                    )
                    result.inventory.append(ref)
                    result.events.append(
                        StageEvent(
                            stage="archive_inspection",
                            status="archive_limit_reached",
                            message=limit_violation,
                            source_id=sid,
                            member_path=info.filename,
                            container_chain=tuple(container_chain),
                        )
                    )
                    continue

                try:
                    # Paso 7: descomprimir el miembro en memoria.
                    member_bytes = zf.read(info.filename)
                except Exception as exc:
                    sid = self._alloc_id()
                    ref = SourceRef(
                        source_id=sid,
                        source_path=source_path,
                        container_chain=tuple(container_chain),
                        member_path=info.filename,
                        nested_depth=member_depth,
                        detected_format=Path(info.filename).suffix[1:].lower() or "unknown",
                        inventory_status="archive_error",
                        status_reason=f"Error descomprimiendo miembro ZIP: {exc}",
                        action_taken="error_recorded",
                        byte_size=info.file_size,
                        compressed_size=info.compress_size,
                    )
                    result.inventory.append(ref)
                    result.work_units.append(
                        WorkUnit(source_ref=ref, error=f"Error descomprimiendo {info.filename}: {exc}")
                    )
                    continue

                # Paso 8: reconocer el formato del miembro por su firma. Si es otro ZIP, se abre con esta misma función;
                # si es un PDF o una imagen, pasa a leerse; si no, se anota como no admitido.
                self._total_uncompressed_bytes += len(member_bytes)
                detected_fmt = detect_format_from_bytes(member_bytes[:64], info.filename)
                member_suffix = Path(info.filename).suffix.lower()

                if (
                    self.config.inspect_nested_archives
                    and (detected_fmt == "zip" or member_suffix in self.config.archive_extensions)
                    and member_suffix not in OFFICE_OPENXML_EXTENSIONS
                ):
                    self._inspect_zip_bytes(
                        raw_bytes=member_bytes,
                        source_path=source_path,
                        container_chain=[*container_chain, info.filename],
                        member_path=info.filename,
                        nested_depth=member_depth,
                        byte_size=len(member_bytes),
                        compressed_size=info.compress_size,
                        result=result,
                        parent_summaries=active_summaries,
                    )
                    continue

                # Documento por leer: se registra en el inventario y se añade a la lista de trabajo.
                if self._is_supported_format(detected_fmt):
                    sid = self._alloc_id()
                    for anc in active_summaries:
                        anc.supported_members += 1
                    ref = SourceRef(
                        source_id=sid,
                        source_path=source_path,
                        container_chain=tuple(container_chain),
                        member_path=info.filename,
                        nested_depth=member_depth,
                        detected_format=detected_fmt,
                        inventory_status="supported",
                        status_reason=f"Miembro ZIP compatible ({detected_fmt}) detectado por firma/contenido",
                        action_taken="processed",
                        byte_size=len(member_bytes),
                        compressed_size=info.compress_size,
                    )
                    result.inventory.append(ref)
                    result.work_units.append(WorkUnit(source_ref=ref, content_bytes=member_bytes))
                    result.supported_unit_count += 1
                else:
                    is_xml = detected_fmt == "xml" or member_suffix == ".xml"
                    for anc in active_summaries:
                        anc.unsupported_members += 1
                        if is_xml:
                            anc.excluded_xml_members += 1
                    self._record_unsupported(
                        result,
                        source_path=source_path,
                        container_chain=tuple(container_chain),
                        member_path=info.filename,
                        nested_depth=member_depth,
                        detected_format="xml" if is_xml else detected_fmt,
                        byte_size=len(member_bytes),
                        compressed_size=info.compress_size,
                        reason=(
                            "Miembro XML dentro de ZIP excluido por formato del flujo visual de comprobantes"
                            if is_xml
                            else f"Miembro ZIP en formato no admitido ({detected_fmt})"
                        ),
                        started=started,
                    )

            result.events.append(
                StageEvent(
                    stage="archive_inspection",
                    status="archive_limit_reached" if summary.limit_reached else "succeeded",
                    message=(
                        f"ZIP inspeccionado ({summary.direct_members} miembros directos, "
                        f"{summary.supported_members} compatibles, estado={summary.status})"
                    ),
                    duration_ms=(time.perf_counter() - started) * 1000.0,
                    source_id=container_sid,
                    member_path=member_path,
                    container_chain=tuple(container_chain),
                )
            )
