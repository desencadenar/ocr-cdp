from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


# Tipos de comprobante que el programa sabe reconocer, con su código oficial de SUNAT.
DEFAULT_DOCUMENT_TYPES: dict[str, str] = {
    "01": "Factura",
    "02": "Recibo por honorarios",
    "03": "Boleta de venta",
    "04": "Liquidación de compra",
    "05": "Boleto de transporte aéreo",
    "07": "Nota de crédito",
    "08": "Nota de débito",
    "12": "Ticket de máquina registradora",
}

# Estados posibles al decidir si un bloque de página es un comprobante de pago (CDP).
# Algunos nombres están repetidos (DETECTED y CDP_DETECTED, por ejemplo) porque el código antiguo
# usaba unos y el nuevo otros; valen exactamente lo mismo.
DETECTED = "cdp_detectado"
CDP_DETECTED = "cdp_detectado"
NOT_RECEIPT = "no_cdp"
NO_CDP = "no_cdp"
AMBIGUOUS = "indeterminado"
UNDETERMINED = "indeterminado"
READ_ERROR = "error_de_lectura"

# Estados de la identidad de un CDP ya detectado: ¿tenemos completos el emisor, el tipo, la serie y el correlativo?
IDENTITY_COMPLETE = "completa"
IDENTITY_INCOMPLETE = "incompleta"
IDENTITY_CONFLICT = "en_conflicto"
IDENTITY_NOT_APPLICABLE = "no_aplica"

# Nombres antiguos de los estados -> nombres actuales.
# Sirve para poder leer bases guardadas con versiones anteriores del programa.
_LEGACY_DETECTION_STATUS_MAP: dict[str, str] = {
    "comprobante_detectado": DETECTED,
    "cdp_detectado": DETECTED,
    "no_comprobante": NOT_RECEIPT,
    "no_cdp": NOT_RECEIPT,
    "ambiguo_o_ilegible": AMBIGUOUS,
    "indeterminado": AMBIGUOUS,
    "error_lectura": READ_ERROR,
    "error_de_lectura": READ_ERROR,
}


# Traduce cualquier estado (antiguo o actual) a uno de los cuatro actuales.
# Si hubo error de lectura manda ese; si no hay ningún estado, queda como indeterminado.
def normalize_detection_status(status: str | None, *, read_error: bool = False) -> str:
    """Normaliza estados de detección históricos o actuales hacia los cuatro estados canónicos."""
    if read_error:
        return READ_ERROR
    if not status:
        return AMBIGUOUS
    return _LEGACY_DETECTION_STATUS_MAP.get(status.strip().lower(), status)




# Resultado de validar un RUC: si tiene la forma correcta, si el dígito de control cuadra
# y su estado en el registro (que este programa nunca consulta).
@dataclass(frozen=True)
class RucValidation:
    value: str
    format_valid: bool
    checksum_valid: bool | None
    registry_status: str = "not_checked"
    reason: str | None = None


# Resultado de normalizar un dato (por ejemplo, el correlativo): el texto original, el normalizado,
# si es válido y, si no lo es, por qué.
@dataclass(frozen=True)
class NormalizationResult:
    original: str | None
    normalized: str | None
    valid: bool
    reason: str | None = None


# Un evento del proceso (etapa, estado, mensaje, página, duración...).
# Se guardan en la base para poder reconstruir qué pasó con cada archivo.
@dataclass(frozen=True)
class StageEvent:
    stage: str
    status: str
    message: str = ""
    page_number: int | None = None
    duration_ms: float | None = None
    source_id: int | None = None
    member_path: str | None = None
    container_chain: tuple[str, ...] = ()
    details: dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "stage": self.stage,
            "status": self.status,
            "message": self.message,
            "page_number": self.page_number,
        }
        if self.duration_ms is not None:
            payload["duration_ms"] = round(self.duration_ms, 3)
        if self.source_id is not None:
            payload["source_id"] = self.source_id
        if self.member_path is not None:
            payload["member_path"] = self.member_path
        if self.container_chain:
            payload["container_chain"] = list(self.container_chain)
        if self.details:
            payload["details"] = dict(self.details)
        return payload


# Una línea de texto de la página junto con su posición (bbox), el bloque al que pertenece
# y la confianza del OCR. Es lo que permite volver al lugar exacto del original.
@dataclass(frozen=True)
class LayoutLine:
    line_number: int
    text: str
    block_number: int = 1
    bbox: tuple[float, float, float, float] | None = None
    ocr_confidence: float | None = None
    source: str = "native_pdf_text"

    def as_dict(self) -> dict[str, Any]:
        return {
            "line_number": self.line_number,
            "text": self.text,
            "block_number": self.block_number,
            "bbox": list(self.bbox) if self.bbox is not None else None,
            "ocr_confidence": self.ocr_confidence,
            "source": self.source,
        }


# Un archivo de la entrada (o un miembro dentro de un ZIP): dónde está, en qué formato viene,
# si se admitió y qué se hizo con él.
@dataclass(frozen=True)
class SourceRef:
    source_id: int
    source_path: str
    container_chain: tuple[str, ...] = ()
    member_path: str | None = None
    nested_depth: int = 0
    detected_format: str = "unknown"
    inventory_status: str = "supported"
    status_reason: str = ""
    action_taken: str = "processed"
    byte_size: int = 0
    compressed_size: int | None = None

    # Texto que dice dónde está el archivo, incluyendo la cadena de ZIP que hubo que abrir para llegar a él.
    @property
    def display_location(self) -> str:
        if not self.container_chain and not self.member_path:
            return self.source_path
        chain = " -> ".join(self.container_chain)
        if self.member_path and (not self.container_chain or self.container_chain[-1] != self.member_path):
            return f"{self.source_path} :: {chain} -> {self.member_path}" if chain else f"{self.source_path} :: {self.member_path}"
        return f"{self.source_path} :: {chain}"

    def as_dict(self) -> dict[str, Any]:
        return {
            "source_id": self.source_id,
            "source_path": self.source_path,
            "container_chain": list(self.container_chain),
            "member_path": self.member_path,
            "nested_depth": self.nested_depth,
            "detected_format": self.detected_format,
            "inventory_status": self.inventory_status,
            "status_reason": self.status_reason,
            "action_taken": self.action_taken,
            "byte_size": self.byte_size,
            "compressed_size": self.compressed_size,
            "display_location": self.display_location,
        }


# Resumen de un ZIP: cuántos miembros tenía, cuántos se pudieron procesar
# y qué incidencias hubo (cifrados, corruptos, límites de seguridad).
@dataclass
class ArchiveSummary:
    source_id: int
    source_path: str
    container_chain: list[str] = field(default_factory=list)
    nested_depth: int = 0
    direct_members: int = 0
    recursive_members: int = 0
    directory_entries: int = 0
    supported_members: int = 0
    unsupported_members: int = 0
    excluded_xml_members: int = 0
    encrypted_members: int = 0
    traversal_blocked_members: int = 0
    corrupted: bool = False
    error_reason: str | None = None
    limit_reached: bool = False
    limit_reasons: list[str] = field(default_factory=list)
    pending_members: list[str] = field(default_factory=list)
    processed_members: int = 0
    unprocessed_members_count: int = 0
    cancelled: bool = False
    max_nested_depth_observed: int = 1
    detected_receipts: int = 0
    ambiguous_records: int = 0
    non_receipts: int = 0
    read_errors: int = 0

    # ¿Se pudo inspeccionar el contenido del ZIP completo?
    @property
    def inspection_status(self) -> str:
        """Distingue si la inspección estructural del contenedor ZIP se completó o tuvo límites/corrupción."""
        if self.corrupted:
            return "corrupto"
        if self.limit_reached:
            return "parcial_limite_alcanzado"
        if self.encrypted_members > 0 or self.traversal_blocked_members > 0:
            return "con_incidencias"
        return "completo"

    # ¿Llegaron a procesarse los miembros que sí se podían procesar?
    @property
    def processing_status(self) -> str:
        """Distingue cuántos miembros elegibles del ZIP llegaron realmente a procesarse."""
        if self.corrupted:
            return "corrupto"
        if self.cancelled or self.unprocessed_members_count > 0:
            return "parcial_cancelado"
        if self.limit_reached:
            return "parcial_limite_alcanzado"
        if self.supported_members == 0:
            return "sin_miembros_procesables"
        if self.read_errors > 0:
            return "con_incidencias"
        return "completo"

    # Estado general del ZIP, mezclando la inspección y el procesamiento.
    @property
    def status(self) -> str:
        if self.corrupted:
            return "corrupto"
        if self.cancelled or self.unprocessed_members_count > 0:
            return "parcial_cancelado"
        if self.limit_reached or self.pending_members:
            return "parcial_limite_alcanzado"
        if self.encrypted_members > 0 or self.traversal_blocked_members > 0 or self.read_errors > 0:
            return "con_incidencias"
        return "completo"

    def as_dict(self) -> dict[str, Any]:
        return {
            "source_id": self.source_id,
            "source_path": self.source_path,
            "container_chain": list(self.container_chain),
            "nested_depth": self.nested_depth,
            "status": self.status,
            "inspection_status": self.inspection_status,
            "processing_status": self.processing_status,
            "direct_members": self.direct_members,
            "recursive_members": self.recursive_members,
            "directory_entries": self.directory_entries,
            "supported_members": self.supported_members,
            "processed_members": self.processed_members,
            "unprocessed_members_count": self.unprocessed_members_count,
            "unsupported_members": self.unsupported_members,
            "excluded_xml_members": self.excluded_xml_members,
            "encrypted_members": self.encrypted_members,
            "traversal_blocked_members": self.traversal_blocked_members,
            "corrupted": self.corrupted,
            "error_reason": self.error_reason,
            "limit_reached": self.limit_reached,
            "limit_reasons": list(self.limit_reasons),
            "pending_members": list(self.pending_members),
            "cancelled": self.cancelled,
            "max_nested_depth_observed": self.max_nested_depth_observed,
            "detected_receipts": self.detected_receipts,
            "ambiguous_records": self.ambiguous_records,
            "non_receipts": self.non_receipts,
            "read_errors": self.read_errors,
        }


# Por cada dato que se lee guardamos de dónde salió: valor observado, valor normalizado,
# página, fragmento, confianza y posición. Es la evidencia que luego se muestra en la hoja de revisión.
@dataclass(frozen=True)
class FieldEvidence:
    observed_value: str
    normalized_value: str | None
    page_number: int
    excerpt: str
    confidence: float
    block_number: int = 1
    line_number: int | None = None
    bbox: tuple[float, float, float, float] | None = None
    source: str = "native_pdf_text"
    ocr_confidence: float | None = None
    heuristic_confidence: float | None = None
    field_status: str = "found"
    decision_comment: str = ""

    def as_dict(self) -> dict[str, object]:
        return {
            "observed_value": self.observed_value,
            "normalized_value": self.normalized_value,
            "page_number": self.page_number,
            "excerpt": self.excerpt,
            "confidence": self.confidence,
            "block_number": self.block_number,
            "line_number": self.line_number,
            "bbox": list(self.bbox) if self.bbox is not None else None,
            "source": self.source,
            "ocr_confidence": self.ocr_confidence,
            "heuristic_confidence": self.heuristic_confidence if self.heuristic_confidence is not None else self.confidence,
            "field_status": self.field_status,
            "decision_comment": self.decision_comment,
        }


# Una fila del detalle de la factura (cantidad, unidad, descripción, precio y total de la línea),
# con el texto observado y su versión normalizada.
@dataclass
class LineItem:
    item_index: int
    code: str | None = None
    description_observed: str | None = None
    description_normalized: str | None = None
    quantity_observed: str | None = None
    quantity_normalized: str | None = None
    unit_observed: str | None = None
    unit_normalized: str | None = None
    unit_price_observed: str | None = None
    unit_price_normalized: str | None = None
    line_total_observed: str | None = None
    line_total_normalized: str | None = None
    page_number: int = 1
    block_number: int = 1
    line_number: int | None = None
    bbox: tuple[float, float, float, float] | None = None
    excerpt: str = ""
    source: str = "native_pdf_text"
    ocr_confidence: float | None = None
    heuristic_confidence: float = 0.88
    missing_reason: str | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "item_index": self.item_index,
            "code": self.code,
            "description_observed": self.description_observed,
            "description_normalized": self.description_normalized,
            "quantity_observed": self.quantity_observed,
            "quantity_normalized": self.quantity_normalized,
            "unit_observed": self.unit_observed,
            "unit_normalized": self.unit_normalized,
            "unit_price_observed": self.unit_price_observed,
            "unit_price_normalized": self.unit_price_normalized,
            "line_total_observed": self.line_total_observed,
            "line_total_normalized": self.line_total_normalized,
            "page_number": self.page_number,
            "block_number": self.block_number,
            "line_number": self.line_number,
            "bbox": list(self.bbox) if self.bbox is not None else None,
            "excerpt": self.excerpt,
            "source": self.source,
            "ocr_confidence": self.ocr_confidence,
            "heuristic_confidence": self.heuristic_confidence,
            "missing_reason": self.missing_reason,
        }


# Todo lo que se sacó de un bloque de página: datos del comprobante, decisiones, evidencias y observaciones.
# Es el objeto que viaja por todo el programa, desde la extracción hasta el Excel.
@dataclass
class ExtractionResult:
    # --- Identificación del comprobante ---
    emitter_ruc: str | None = None
    recipient_ruc: str | None = None
    document_type: str | None = None
    series: str | None = None
    correlativo_original: str | None = None
    correlativo_normalized: str | None = None
    # --- Fechas, nombres y moneda ---
    issue_date: str | None = None
    due_date: str | None = None
    issuer_name: str | None = None
    recipient_name: str | None = None
    currency: str | None = None
    # --- Importes tal como se leyeron (texto) ---
    subtotal: str | None = None
    valor_venta: str | None = None
    gravado: str | None = None
    inafecto: str | None = None
    exonerado: str | None = None
    no_gravado: str | None = None
    gratuitas: str | None = None
    exportaciones: str | None = None
    descuentos: str | None = None
    otros_cargos: str | None = None
    otros_tributos: str | None = None
    isc: str | None = None
    icbper: str | None = None
    recargo_consumo: str | None = None
    recargo_fise: str | None = None
    redondeo: str | None = None
    anticipos: str | None = None
    importe_detraccion: str | None = None
    importe_retencion: str | None = None
    importe_percepcion: str | None = None
    igv: str | None = None
    total: str | None = None
    total_a_pagar: str | None = None
    # --- Otros datos ---
    payment_condition: str | None = None
    referenced_document: str | None = None
    general_description: str | None = None
    # --- Decisiones del programa: ¿es un comprobante?, ¿está completa su identidad? ---
    detection_status: str = AMBIGUOUS
    identity_status: str = IDENTITY_NOT_APPLICABLE
    comprobante_detectado: bool = False
    identidad_completa: bool = False
    renombrable: bool = False
    # --- Dónde está en el archivo (páginas y bloques) ---
    page_start: int = 1
    page_end: int = 1
    block_start: int = 1
    block_end: int = 1
    reading_source: str = "native_pdf_text"
    # --- Por qué se decidió así ---
    classification_rule: str = "unresolved"
    non_receipt_type: str | None = None
    decision_comment: str = ""
    detection_reason: str = ""
    detection_excerpt: str = ""
    identity_reason: str = ""
    identity_excerpt: str = ""
    # --- Roles de los RUC (emisor, receptor y terceros) ---
    recipient_ruc_match: bool = False
    recipient_role_status: str = "not_found"
    issuer_role_status: str = "not_found"
    third_party_rucs: list[str] = field(default_factory=list)
    candidate_fiscal_key: str | None = None
    secondary_review_fields: list[str] = field(default_factory=list)
    requires_human_review: bool = True
    # --- Detalle, controles, evidencias y observaciones ---
    line_items: list[LineItem] = field(default_factory=list)
    derived_checks: list[dict[str, Any]] = field(default_factory=list)
    field_statuses: dict[str, str] = field(default_factory=dict)
    fields_original: dict[str, str | None] = field(default_factory=dict)
    confidence: dict[str, float] = field(default_factory=dict)
    field_evidence: dict[str, list[FieldEvidence]] = field(default_factory=dict)
    evidence: dict[str, list[str]] = field(default_factory=dict)
    candidates: dict[str, list[str]] = field(default_factory=dict)
    observations: list[str] = field(default_factory=list)
    # --- Errores de lectura y texto completo del bloque ---
    read_error: bool = False
    read_error_details: list[str] = field(default_factory=list)
    extracted_text: str = ""

    # Alias para el código antiguo: sólo refleja el estado de detección.
    @property
    def status(self) -> str:
        """Alias para el código antiguo; sólo refleja el estado de detección."""
        return self.detection_status

    # Al asignarlo se normaliza el estado y se actualiza si el bloque cuenta como comprobante detectado.
    @status.setter
    def status(self, value: str) -> None:
        normalized = normalize_detection_status(value, read_error=self.read_error)
        self.detection_status = normalized
        self.comprobante_detectado = normalized == DETECTED

    # Los datos ya normalizados, en un diccionario plano. Es lo que se guarda y lo que se exporta.
    @property
    def fields_normalized(self) -> dict[str, str | None]:
        return {
            "emitter_ruc": self.emitter_ruc,
            "recipient_ruc": self.recipient_ruc,
            "document_type": self.document_type,
            "series": self.series,
            "correlativo_original": self.correlativo_original,
            "correlativo_normalized": self.correlativo_normalized,
            "issue_date": self.issue_date,
            "due_date": self.due_date,
            "issuer_name": self.issuer_name,
            "recipient_name": self.recipient_name,
            "currency": self.currency,
            "subtotal": self.subtotal,
            "valor_venta": self.valor_venta,
            "gravado": self.gravado,
            "inafecto": self.inafecto,
            "exonerado": self.exonerado,
            "no_gravado": self.no_gravado,
            "gratuitas": self.gratuitas,
            "exportaciones": self.exportaciones,
            "descuentos": self.descuentos,
            "otros_cargos": self.otros_cargos,
            "otros_tributos": self.otros_tributos,
            "isc": self.isc,
            "icbper": self.icbper,
            "recargo_consumo": self.recargo_consumo,
            "recargo_fise": self.recargo_fise,
            "redondeo": self.redondeo,
            "anticipos": self.anticipos,
            "importe_detraccion": self.importe_detraccion,
            "importe_retencion": self.importe_retencion,
            "importe_percepcion": self.importe_percepcion,
            "igv": self.igv,
            "total": self.total,
            "total_a_pagar": self.total_a_pagar,
            "payment_condition": self.payment_condition,
            "referenced_document": self.referenced_document,
            "general_description": self.general_description,
            "third_party_rucs": ",".join(self.third_party_rucs) if self.third_party_rucs else None,
            "candidate_fiscal_key": self.candidate_fiscal_key,
            "identity_status": self.identity_status,
        }

    # DEPRECADO: ningún módulo lo usa. Armaba un nombre de archivo a partir de la identidad.
    @property
    def normalized_stem(self) -> str | None:
        if not self.renombrable or not self.identidad_completa:
            return None
        return "-".join(
            (
                self.emitter_ruc or "",
                self.document_type or "",
                self.series or "",
                self.correlativo_normalized or "",
            )
        )

    # Clave que identifica un comprobante: emisor, tipo, serie y correlativo.
    # Dos documentos con la misma clave se consideran el mismo comprobante.
    @property
    def identity_key(self) -> tuple[str, str, str, str] | None:
        if not self.identidad_completa:
            return None
        return (
            self.emitter_ruc or "",
            self.document_type or "",
            (self.series or "").casefold(),
            self.correlativo_normalized or "",
        )

    # DEPRECADO: ningún módulo lo usa. Armaba una clave con lo observado aunque la identidad no estuviera completa.
    @property
    def observed_key(self) -> str | None:
        corr = self.correlativo_normalized or self.correlativo_original
        if not (self.series and corr):
            return None
        emisor = self.emitter_ruc or (
            self.candidates.get("emitter_ruc", [None])[0]
            if len(self.candidates.get("emitter_ruc", [])) == 1
            else None
        ) or "RUC_PENDIENTE"
        doc_type = self.document_type or "TIPO_PENDIENTE"
        return f"{emisor}-{doc_type}-{self.series.upper()}-{corr}"

    # Convierte el resultado a un diccionario para guardarlo en la base.
    def as_dict(self) -> dict[str, object]:
        return {
            "detection_status": self.detection_status,
            "identity_status": self.identity_status,
            "comprobante_detectado": self.comprobante_detectado,
            "identidad_completa": self.identidad_completa,
            "renombrable": self.renombrable,
            "page_start": self.page_start,
            "page_end": self.page_end,
            "block_start": self.block_start,
            "block_end": self.block_end,
            "reading_source": self.reading_source,
            "classification_rule": self.classification_rule,
            "non_receipt_type": self.non_receipt_type,
            "decision_comment": self.decision_comment,
            "detection_reason": self.detection_reason,
            "detection_excerpt": self.detection_excerpt,
            "identity_reason": self.identity_reason,
            "identity_excerpt": self.identity_excerpt,
            "recipient_ruc_match": self.recipient_ruc_match,
            "recipient_role_status": self.recipient_role_status,
            "issuer_role_status": self.issuer_role_status,
            "third_party_rucs": list(self.third_party_rucs),
            "candidate_fiscal_key": self.candidate_fiscal_key,
            "secondary_review_fields": list(self.secondary_review_fields),
            "general_description": self.general_description,
            "requires_human_review": self.requires_human_review,
            "line_items": [item.as_dict() for item in self.line_items],
            "derived_checks": list(self.derived_checks),
            "field_statuses": dict(self.field_statuses),
            "fields_original": dict(self.fields_original),
            "fields_normalized": self.fields_normalized,
            "field_evidence": {
                name: [item.as_dict() for item in items]
                for name, items in self.field_evidence.items()
            },
            "confidence": dict(self.confidence),
            "candidates": {name: list(items) for name, items in self.candidates.items()},
            "observations": list(self.observations),
            "read_error": self.read_error,
            "read_error_details": list(self.read_error_details),
        }
