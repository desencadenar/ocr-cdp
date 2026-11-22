"""Pruebas automáticas con documentos sintéticos para verificar extracción, cuadres tributarios, padrón, expediente y exportación."""

from __future__ import annotations

import contextlib
from dataclasses import replace
from decimal import Decimal
import io
import json
from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest
import zipfile

import openpyxl
from PIL import Image
import pymupdf

ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from ocr_fiscal.amounts import amount_mentions, evaluate_igv_rate, reconcile_amounts
from ocr_fiscal.cli import main as cli_main
from ocr_fiscal.config import AppConfig, load_config
from ocr_fiscal.documents import PageRead
from ocr_fiscal.exports import export_results
from ocr_fiscal.extraction import parse_document
from ocr_fiscal.models import (
    AMBIGUOUS,
    DEFAULT_DOCUMENT_TYPES,
    DETECTED,
    IDENTITY_COMPLETE,
    LayoutLine,
    NOT_RECEIPT,
    ExtractionResult,
)
from ocr_fiscal.ocr import OCRResult
from ocr_fiscal.padron import (
    PADRON_AUSENTE,
    PADRON_PRESENTE,
    LocalPadronAdapter,
    cargar_padron_staging,
    compare_names_with_padron,
)
from ocr_fiscal.processor import FiscalProcessor, _crosscheck_sunat_qr
from ocr_fiscal.tax import normalize_correlativo, validate_dni, validate_ruc
from ocr_fiscal.visible import extract_spot_details


def make_valid_ruc(prefix10: str) -> str:
    """Construye un RUC peruano de 11 dígitos con dígito verificador módulo 11 válido."""
    assert len(prefix10) == 10 and prefix10.isdigit()
    weights = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    remainder = sum(int(d) * w for d, w in zip(prefix10, weights)) % 11
    check = 11 - remainder
    if check == 10:
        check = 0
    elif check == 11:
        check = 1
    ruc = f"{prefix10}{check}"
    assert validate_ruc(ruc).checksum_valid is True
    return ruc


def make_page(
    lines: list[str],
    *,
    page_number: int = 1,
    source: str = "native_pdf_text",
    qr_payloads: list[str] | None = None,
) -> PageRead:
    layout = [
        LayoutLine(line_number=i, text=line, block_number=1, bbox=(10.0, float(i * 15), 500.0, float(i * 15 + 12)), source=source)
        for i, line in enumerate(lines, start=1)
    ]
    return PageRead(
        page_number=page_number,
        text="\n".join(lines),
        source=source,
        status="ok",
        quality={"char_count": sum(len(x) for x in lines), "line_count": len(lines)},
        layout_lines=layout,
        qr_payloads=list(qr_payloads or []),
    )


def make_text_pdf(path: Path, pages_lines: list[list[str]]) -> Path:
    doc = pymupdf.open()
    try:
        for lines in pages_lines:
            page = doc.new_page(width=595, height=842)
            y = 40
            for line in lines:
                page.insert_text((36, y), line, fontsize=10, fontname="helv")
                y += 14
        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(path)
    finally:
        doc.close()
    return path


def make_padron_sqlite(db_path: Path, rows: list[tuple[str, str, str, str, str]], cutoff_date: str = "2026-03-01") -> Path:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with contextlib.closing(sqlite3.connect(db_path)) as conn:
        with conn:
            conn.executescript(
                """
                DROP TABLE IF EXISTS padron;
                CREATE TABLE padron (
                    ruc TEXT PRIMARY KEY,
                    razon_social TEXT NOT NULL,
                    estado TEXT NOT NULL,
                    cond_domicilio TEXT NOT NULL,
                    ubigeo TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS padron_metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                """
            )
            conn.executemany(
                "INSERT INTO padron (ruc, razon_social, estado, cond_domicilio, ubigeo) VALUES (?,?,?,?,?)",
                rows,
            )
            conn.executemany(
                "INSERT OR REPLACE INTO padron_metadata (key, value) VALUES (?,?)",
                [("cutoff_date", cutoff_date), ("source_path", "test_padron.txt")],
            )
            conn.commit()
    return db_path


class MotorOCRFicticio:
    """Motor OCR de prueba que devuelve líneas predefinidas sin invocar Tesseract."""

    name = "motor_ocr_ficticio"
    version = "1.0"

    def __init__(self, text_by_hint: str = "") -> None:
        self.text_by_hint = text_by_hint

    def availability(self) -> tuple[bool, str | None]:
        return True, None

    def recognize(self, image_path: str | Path) -> OCRResult:
        lines = [line for line in self.text_by_hint.splitlines() if line.strip()]
        layout = [
            LayoutLine(line_number=i, text=t, block_number=1, bbox=(10.0, float(i * 15), 500.0, float(i * 15 + 12)), ocr_confidence=0.95, source="ocr")
            for i, t in enumerate(lines, start=1)
        ]
        return OCRResult(
            text=self.text_by_hint,
            mean_confidence=0.95,
            engine_version="1.0",
            layout_lines=layout,
        )


# RUC completamente ficticios (prefijo 2099 / 1099) con dígito de control válido sólo para pruebas sintéticas.
RUC_EMISOR_A = make_valid_ruc("2099000033")
RUC_EMISOR_B = make_valid_ruc("2099000044")
RUC_NAT_C = make_valid_ruc("1099000055")
RUC_RECEPTOR = make_valid_ruc("2099000011")
RUC_OTRO = make_valid_ruc("2099000022")


class TestTaxAndConfig(unittest.TestCase):
    def test_document_types_and_dni_correlativo(self) -> None:
        for code in ("01", "02", "03", "04", "05", "07", "08", "10", "12", "13", "14", "16", "20", "40", "50", "52", "54"):
            self.assertIn(code, DEFAULT_DOCUMENT_TYPES)

        self.assertTrue(validate_dni("44556677"))
        self.assertFalse(validate_dni("4455667"))
        self.assertFalse(validate_dni("4455667A"))

        std_ok = normalize_correlativo("169", document_type="01")
        self.assertTrue(std_ok.valid)
        self.assertEqual(std_ok.normalized, "00000169")

        std_too_long = normalize_correlativo("123456789", document_type="01")
        self.assertFalse(std_too_long.valid)

        for ext_type in ("05", "10", "12", "13", "14", "16", "50", "52", "54"):
            ext = normalize_correlativo("001234567890", document_type=ext_type)
            self.assertTrue(ext.valid, f"Tipo {ext_type} debería admitir correlativo largo")
            self.assertEqual(ext.normalized, "001234567890")

    def test_load_config_new_keys(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            cfg_path = Path(tmp) / "config.json"
            cfg_path.write_text(
                json.dumps(
                    {
                        "unificar_notas_en_proveedor": False,
                        "tolerancia_redondeo": 0.05,
                        "padron_dias_maximos": 45,
                    }
                ),
                encoding="utf-8",
            )
            cfg = load_config(cfg_path)
            self.assertFalse(cfg.unify_notes_by_provider)
            self.assertAlmostEqual(cfg.rounding_tolerance, 0.05)
            self.assertEqual(cfg.padron_max_age_days, 45)


class TestAmountsAndIGV(unittest.TestCase):
    def test_amount_mentions_clean_parsing_and_no_contamination(self) -> None:
        # "Descuento Total" no debe contaminar "total"
        mentions_desc = list(amount_mentions("Descuento Total: S/ 15.00"))
        fields_desc = {m[0] for m in mentions_desc}
        self.assertIn("descuentos", fields_desc)
        self.assertNotIn("total", fields_desc)

        # S7 / SY reconocidos como PEN por ruido OCR
        m_s7 = list(amount_mentions("Importe Total S7 118.00"))
        self.assertEqual(m_s7[-1][2], "PEN")
        m_sy = list(amount_mentions("Importe Total SY 118.00"))
        self.assertEqual(m_sy[-1][2], "PEN")

        # Retención entre paréntesis en RHE -> negativa
        m_ret = list(amount_mentions("Retención (8%) IR: (160.00)"))
        ret = [m for m in m_ret if m[0] == "importe_retencion"]
        self.assertTrue(ret)
        self.assertEqual(Decimal(ret[0][1]), Decimal("-160.00"))

        # "IGV 18%" en etiqueta de base gravada no debe capturarse como importe IGV
        m_base = list(amount_mentions("Base Imponible IGV 18%: 100.00"))
        self.assertTrue(any(m[0] == "gravado" and Decimal(m[1]) == Decimal("100.00") for m in m_base))
        self.assertFalse(any(m[0] == "igv" and Decimal(m[1]) == Decimal("18.00") for m in m_base))

        # "Detracción SPOT 12%" captura detracción sin contaminar IGV con el porcentaje
        m_spot = list(amount_mentions("Detracción SPOT 12%: 12.00 | Base IGV: 100.00"))
        self.assertTrue(any(m[0] == "importe_detraccion" and Decimal(m[1]) == Decimal("12.00") for m in m_spot))

    def test_reconcile_amounts_peruvian_tax_cases(self) -> None:
        # 1. 100% gratuita
        r_grat = ExtractionResult(document_type="01", gratuitas="150.00", total="0.00")
        c_grat = reconcile_amounts(r_grat)
        self.assertTrue(c_grat["consistent"])
        self.assertEqual(c_grat["status"], "cuadra")

        # 2. Gratuita con IGV referencial
        r_grat_igv = ExtractionResult(document_type="01", gratuitas="100.00", igv="18.00", total="0.00")
        c_grat_igv = reconcile_amounts(r_grat_igv)
        self.assertTrue(c_grat_igv["consistent"])

        # 3. Percepción en total_a_pagar
        r_perc = ExtractionResult(
            document_type="01",
            gravado="100.00",
            igv="18.00",
            total="118.00",
            importe_percepcion="2.36",
            total_a_pagar="120.36",
        )
        c_perc = reconcile_amounts(r_perc)
        self.assertTrue(c_perc["consistent"])
        self.assertEqual(c_perc["status"], "cuadra")

        # 4. RHE con retención 8%
        r_rhe = ExtractionResult(
            document_type="02",
            total="2000.00",
            importe_retencion="160.00",
            total_a_pagar="1840.00",
        )
        c_rhe = reconcile_amounts(r_rhe)
        self.assertTrue(c_rhe["consistent"])

        # 5. ISC dentro de base gravada
        r_isc = ExtractionResult(document_type="01", gravado="100.00", isc="20.00", igv="21.60", total="141.60")
        c_isc = reconcile_amounts(r_isc)
        self.assertTrue(c_isc["consistent"])

        # 6. Tolerancia redondeo ±0.02
        r_round = ExtractionResult(document_type="01", gravado="100.00", igv="18.01", total="118.00")
        c_round = reconcile_amounts(r_round, rounding_tolerance=Decimal("0.02"))
        self.assertTrue(c_round["consistent"])
        self.assertEqual(c_round["status"], "cuadra_con_redondeo")

        # 7. Nota de crédito: signos coherentes cuadran; signos mezclados entre componentes van a revisión
        #.
        r_nc_pos = ExtractionResult(document_type="07", gravado="100.00", igv="18.00", total="-118.00")
        self.assertTrue(reconcile_amounts(r_nc_pos)["consistent"])
        r_nc_mix = ExtractionResult(document_type="07", gravado="-100.00", igv="18.00", total="-118.00")
        self.assertFalse(reconcile_amounts(r_nc_mix)["consistent"])

        r_nc_13 = ExtractionResult(document_type="07", note_type_code="13", total="0.00")
        self.assertTrue(reconcile_amounts(r_nc_13)["consistent"])

        # 8. FISE ya contenido en inafecto no genera falso descuadre
        r_fise = ExtractionResult(
            document_type="14",
            gravado="100.00",
            inafecto="12.00",
            recargo_fise="2.00",
            igv="18.00",
            total="130.00",
        )
        self.assertTrue(reconcile_amounts(r_fise)["consistent"])

    def test_evaluate_igv_rate_variable_rates(self) -> None:
        # 18% estándar
        chk_18 = evaluate_igv_rate(ExtractionResult(document_type="01", issue_date="2026-03-01", gravado="100.00", igv="18.00", total="118.00"))
        self.assertTrue(chk_18["consistent"])
        self.assertEqual(chk_18["matched_rate"], 18.0)

        # 10 % régimen especial restaurantes y hoteles en 2025 (8 % IGV Ley 31556 + 2 % IPM)
        chk_10 = evaluate_igv_rate(ExtractionResult(document_type="01", issue_date="2025-08-15", gravado="100.00", igv="10.00", total="110.00"))
        # El giro del emisor no se conoce, así que la tasa especial
        # se reconoce pero queda para verificar (antes se exigía consistent True).
        self.assertEqual(chk_10["status"], "tasa_especial_verificar")
        self.assertTrue(chk_10["needs_review"])
        self.assertEqual(chk_10["matched_rate"], 10.0)

        # 10.5 % régimen especial restaurantes y hoteles en 2026 (8 % IGV Ley 32219 + 2.5 % IPM Ley 32387)
        chk_105 = evaluate_igv_rate(ExtractionResult(document_type="01", issue_date="2026-04-10", gravado="100.00", igv="10.50", total="110.50"))
        # El giro del emisor no se conoce, así que la tasa especial
        # se reconoce pero queda para verificar (antes se exigía consistent True).
        self.assertEqual(chk_105["status"], "tasa_especial_verificar")
        self.assertTrue(chk_105["needs_review"])
        self.assertEqual(chk_105["matched_rate"], 10.5)

        # 15 % régimen especial restaurantes y hoteles en 2027 (12 % IGV Ley 32219 + 3.0 % IPM Ley 32387)
        chk_15 = evaluate_igv_rate(ExtractionResult(document_type="01", issue_date="2027-02-15", gravado="100.00", igv="15.00", total="115.00"))
        # El giro del emisor no se conoce, así que la tasa especial
        # se reconoce pero queda para verificar (antes se exigía consistent True).
        self.assertEqual(chk_15["status"], "tasa_especial_verificar")
        self.assertTrue(chk_15["needs_review"])
        self.assertEqual(chk_15["matched_rate"], 15.0)

        # 10 % en nota de crédito emitida en 2026 por operación de 2025
        chk_nc_prior = evaluate_igv_rate(ExtractionResult(document_type="07", issue_date="2026-01-15", gravado="100.00", igv="10.00", total="110.00"))
        # El giro del emisor no se conoce, así que la tasa especial
        # se reconoce pero queda para verificar (antes se exigía consistent True).
        self.assertEqual(chk_nc_prior["status"], "tasa_especial_verificar")
        self.assertTrue(chk_nc_prior["needs_review"])
        self.assertEqual(chk_nc_prior["matched_rate"], 10.0)

        # 4% IVAP: solo con mención de IVAP o arroz pilado
        chk_4 = evaluate_igv_rate(ExtractionResult(document_type="01", issue_date="2026-02-01", gravado="100.00", igv="4.00", total="104.00", extracted_text="ARROZ PILADO (IVAP)"))
        self.assertTrue(chk_4["consistent"])
        self.assertEqual(chk_4["matched_rate"], 4.0)

        # 0% exonerado / inafecto
        chk_0 = evaluate_igv_rate(ExtractionResult(document_type="01", issue_date="2026-02-01", exonerado="100.00", igv="0.00", total="100.00"))
        self.assertTrue(chk_0["consistent"])
        self.assertEqual(chk_0["status"], "exonerado_inafecto_sin_igv")
        self.assertEqual(chk_0["effective_rate"], 0.0)

        # Tasa atípica -> observada (consistent=False)
        chk_bad = evaluate_igv_rate(ExtractionResult(document_type="01", issue_date="2026-02-01", gravado="100.00", igv="14.00", total="114.00"))
        self.assertFalse(chk_bad["consistent"])
        self.assertEqual(chk_bad["status"], "tasa_igv_atipica")


class TestExtractionAndMetadata(unittest.TestCase):
    def test_factura_citing_oc_guia_estado_cuenta_not_dropped_as_no_cdp(self) -> None:
        page = make_page(
            [
                "SERVICIOS INDUSTRIALES ANDINOS S.A.C.",
                f"RUC: {RUC_EMISOR_A}",
                "FACTURA ELECTRONICA",
                "F001-00001234",
                "Fecha de Emisión: 15/03/2026",
                "Fecha de Vencimiento: 14/04/2026",
                "Condición de Pago: Crédito",
                "Nro Cuotas: 2",
                f"Señor(es): EMPRESA COMPRADORA S.A. RUC: {RUC_RECEPTOR}",
                "Orden de Compra: OC-2026-0091",
                "Guía de Remisión: T001-00000456",
                "Placa: ABC-123",
                "Referencia Estado de Cuenta: N/A",
                "1.250 KG MANTENIMIENTO BOMBA HIDRAULICA S/ 1,000.00",
                "Observaciones:",
                "Servicio prestado en planta Callao",
                "Operación sujeta al SPOT 12% Cód. 022 CTA BN 00-000-123456",
                "Op. Gravada: S/ 1,000.00",
                "IGV 18%: S/ 180.00",
                "Importe Total: S/ 1,180.00",
            ]
        )
        res_list = parse_document([page], target_recipient_ruc=RUC_RECEPTOR)
        self.assertEqual(len(res_list), 1)
        r = res_list[0]
        self.assertEqual(r.detection_status, DETECTED)
        self.assertEqual(r.identity_status, IDENTITY_COMPLETE)
        self.assertEqual(r.document_type, "01")
        self.assertEqual(r.series, "F001")
        self.assertEqual(r.correlativo_normalized, "00001234")
        self.assertEqual(r.due_date, "2026-04-14")
        self.assertEqual((r.payment_condition or "").lower().replace("é", "e"), "credito")
        self.assertEqual(r.installments_count, "2")
        self.assertIn("OC-2026-0091", r.purchase_order_ref or "")
        self.assertIn("T001-00000456", r.remision_guide_ref or "")
        self.assertEqual(r.vehicle_plate, "ABC-123")
        self.assertEqual(r.detraccion_rate, "12")
        self.assertEqual(r.detraccion_code, "022")
        self.assertEqual(r.detraccion_account, "00-000-123456")
        self.assertIn("planta Callao", r.observations_text or "")

        # La cantidad con 3 decimales ("1.250") no debe convertirse en "1250"
        if r.line_items:
            qty_vals = [it.quantity_normalized for it in r.line_items if it.quantity_normalized]
            self.assertNotIn("1250", qty_vals)
            self.assertIn("1.250", qty_vals)
            # No se debe inventar unit_price si la fila solo trajo cantidad, desc y total
            for it in r.line_items:
                self.assertIsNone(it.unit_price_normalized)

    def test_no_cdp_types_vs_new_cdp_types(self) -> None:
        # La constancia SPOT no es comprobante de pago: se separa como soporte (constancia_detraccion_spot)
        spot_page = make_page(
            [
                "BANCO DE LA NACION - CONSTANCIA DE DEPOSITO",
                "SISTEMA DE PAGO DE OBLIGACIONES TRIBUTARIAS - SPOT",
                f"RUC Proveedor: {RUC_EMISOR_A}",
                "Tipo de Bien / Servicio: 022 Otros servicios",
                "Importe Depósito: S/ 240.00",
                "Número de Constancia: 123456789",
            ]
        )
        res_spot = parse_document([spot_page], target_recipient_ruc=RUC_RECEPTOR)[0]
        self.assertEqual(res_spot.detection_status, NOT_RECEIPT)
        self.assertEqual(res_spot.non_receipt_type, "constancia_detraccion_spot")

        # Estado de cuenta bancario puro -> se descarta como no_cdp (estado_de_cuenta_bancario)
        bank_page = make_page(
            [
                "BANCO EJEMPLO DEMO S.A.",
                "ESTADO DE CUENTA CORRIENTE",
                "Saldo Anterior: 1,500.00",
                "Movimientos del Periodo",
            ]
        )
        res_bank = parse_document([bank_page], target_recipient_ruc=RUC_RECEPTOR)[0]
        self.assertEqual(res_bank.detection_status, NOT_RECEIPT)
        self.assertEqual(res_bank.non_receipt_type, "estado_de_cuenta_bancario")

        # Recibo de servicios públicos (14) que menciona "Estado de cuenta anterior" -> sí es comprobante
        recibo_14 = make_page(
            [
                "DISTRIBUIDORA ELECTRICA DEMO S.A.A.",
                f"RUC: {RUC_EMISOR_B}",
                "RECIBO POR SERVICIOS PUBLICOS",
                "S001-000123456789",
                "Fecha de Emisión: 05/03/2026",
                f"RUC Cliente: {RUC_RECEPTOR}",
                "Estado de cuenta anterior: S/ 0.00",
                "Subtotal: S/ 100.00",
                "IGV 18%: S/ 18.00",
                "Total a Pagar: S/ 118.00",
            ]
        )
        res_14 = parse_document([recibo_14], target_recipient_ruc=RUC_RECEPTOR)[0]
        self.assertEqual(res_14.detection_status, DETECTED)
        self.assertEqual(res_14.document_type, "14")
        self.assertEqual(res_14.correlativo_normalized, "000123456789")

        # Nuevos tipos SUNAT: 10, 13, 16, 20, 40, 50, 05 (aéreo)
        type_cases = [
            ("10", ["RECIBO POR ARRENDAMIENTO", f"RUC: {RUC_NAT_C}", "1683-00123456", f"RUC Arrendatario: {RUC_RECEPTOR}", "Fecha de Emisión: 01/03/2026", "Total: S/ 1,500.00"]),
            ("13", ["DOCUMENTO AUTORIZADO SBS - SEGURO", f"RUC: {RUC_EMISOR_A}", "F001-0000999900", f"RUC: {RUC_RECEPTOR}", "Fecha de Emisión: 02/03/2026", "Op. Gravada: 100.00", "IGV: 18.00", "Importe Total: 118.00"]),
            ("16", ["BOLETO DE VIAJE TERRESTRE", f"RUC: {RUC_EMISOR_A}", "B001-000123456", f"RUC: {RUC_RECEPTOR}", "Fecha de Emisión: 03/03/2026", "Total: S/ 80.00"]),
            ("20", ["COMPROBANTE DE RETENCION ELECTRONICO", f"RUC: {RUC_EMISOR_A}", "R001-00000123", f"RUC: {RUC_RECEPTOR}", "Fecha de Emisión: 04/03/2026", "Importe Retención: S/ 30.00", "Total: S/ 30.00"]),
            ("40", ["COMPROBANTE DE PERCEPCION ELECTRONICO", f"RUC: {RUC_EMISOR_A}", "P001-00000123", f"RUC: {RUC_RECEPTOR}", "Fecha de Emisión: 04/03/2026", "Importe Percepción: S/ 20.00", "Total: S/ 20.00"]),
            ("50", ["DECLARACION ADUANERA DE MERCANCIAS (DUA)", f"RUC: {RUC_EMISOR_A}", "118-2026-10-001234", f"RUC: {RUC_RECEPTOR}", "Fecha de Emisión: 04/03/2026", "Total: USD 500.00"]),
            ("05", ["COMPROBANTE DE COMPRA - BOLETO AEREO", "IATA: 9150012", f"RUC: {RUC_EMISOR_A}", "544-2401234567890", f"RUC: {RUC_RECEPTOR}", "Fecha de Emisión: 05/03/2026", "Total: USD 350.00"]),
        ]
        for expected_type, lines in type_cases:
            res = parse_document([make_page(lines)], target_recipient_ruc=RUC_RECEPTOR)[0]
            self.assertEqual(res.detection_status, DETECTED, f"Falló detección tipo {expected_type}: {res.decision_comment}")
            self.assertEqual(res.document_type, expected_type, f"Tipo esperado {expected_type}, obtenido {res.document_type}")

    def test_boleta_dni_nc_multiref_rhe_retencion_and_currency_conflict(self) -> None:
        # Boleta con DNI 8 dígitos
        boleta = make_page(
            [
                "RESTAURANTE LA COSTA S.A.C.",
                f"RUC: {RUC_EMISOR_A}",
                "BOLETA DE VENTA ELECTRONICA",
                "B001-00000777",
                "Fecha de Emisión: 10/03/2026",
                "Cliente: JUAN PEREZ DNI: 44556677",
                "Consumo de alimentos",
                "Total a Pagar: S/ 50.00",
            ]
        )
        res_bol = parse_document([boleta], target_recipient_ruc=RUC_RECEPTOR)[0]
        self.assertEqual(res_bol.detection_status, DETECTED)
        self.assertIsNone(res_bol.recipient_ruc)
        self.assertEqual(res_bol.recipient_doc_type, "DNI 44556677")
        self.assertEqual(res_bol.recipient_role_status, "verified_dni")

        # Nota de crédito con múltiples documentos afectados y motivo SUNAT
        nc = make_page(
            [
                "PROVEEDOR ALFA S.A.C.",
                f"RUC: {RUC_EMISOR_A}",
                "NOTA DE CREDITO ELECTRONICA",
                "FC01-00000055",
                "Fecha de Emisión: 12/03/2026",
                f"Cliente RUC: {RUC_RECEPTOR}",
                "Tipo de Nota: 01 - Anulación de la operación",
                "Documento que modifica: F001-00000010, F001-00000011",
                "Op. Gravada: S/ 100.00",
                "IGV: S/ 18.00",
                "Importe Total: S/ 118.00",
            ]
        )
        res_nc = parse_document([nc], target_recipient_ruc=RUC_RECEPTOR)[0]
        self.assertEqual(res_nc.document_type, "07")
        self.assertEqual(res_nc.note_type_code, "01")
        self.assertIn("F001-00000010", res_nc.referenced_documents)
        self.assertIn("F001-00000011", res_nc.referenced_documents)

        # RHE mayor a S/ 1,500 sin retención ni suspensión -> dispara control rhe_retencion_check
        rhe = make_page(
            [
                "PEREZ GOMEZ JUAN CARLOS",
                f"RUC: {RUC_NAT_C}",
                "RECIBO POR HONORARIOS ELECTRONICO",
                "E001-00000088",
                "Fecha de Emisión: 10/03/2026",
                f"Recibí de: EMPRESA COMPRADORA RUC: {RUC_RECEPTOR}",
                "Por concepto de: Asesoría tributaria mensual",
                "Total por Honorarios: S/ 2,000.00",
                "Retención (8%) IR: S/ 0.00",
                "Total Neto Recibido: S/ 2,000.00",
            ]
        )
        res_rhe = parse_document([rhe], target_recipient_ruc=RUC_RECEPTOR)[0]
        rhe_chks = [c for c in res_rhe.derived_checks if c.get("check") == "rhe_retencion_check"]
        self.assertTrue(rhe_chks and not rhe_chks[0]["consistent"])

        # Conflicto de moneda (S/ y USD mezclados) no adivina
        mixed_curr = make_page(
            [
                "IMPORTADORA DELTA S.A.C.",
                f"RUC: {RUC_EMISOR_A}",
                "FACTURA ELECTRONICA",
                "F002-00000019",
                "Fecha de Emisión: 10/03/2026",
                f"RUC: {RUC_RECEPTOR}",
                "Op. Gravada: USD 100.00",
                "IGV: S/ 18.00",
                "Importe Total: S/ 118.00",
            ]
        )
        res_mix = parse_document([mixed_curr], target_recipient_ruc=RUC_RECEPTOR)[0]
        self.assertIsNone(res_mix.currency)
        self.assertIn("currency", res_mix.secondary_review_fields)

    def test_extract_spot_details_and_decimal_preserved(self) -> None:
        details = extract_spot_details(["Operación sujeta al SPOT 10% Código 022 Banco de la Nación 00-058-123456"])
        self.assertIsNotNone(details["rate"])
        self.assertIsNotNone(details["code"])
        self.assertIsNotNone(details["account"])
        assert details["rate"] and details["code"] and details["account"]
        self.assertEqual(details["rate"][0], "10")
        self.assertEqual(details["code"][0], "022")
        self.assertEqual(details["account"][0], "00-058-123456")

        # Monto 1.250 (punto seguido de tres dígitos) es ambiguo: puede ser 1 250 o 1.25. No se elige
        # ninguna lectura; el total queda en blanco y el comprobante va a revisión.
        p_dec = make_page(
            [
                "FERRETERIA NORTE S.A.C.",
                f"RUC: {RUC_EMISOR_A}",
                "FACTURA ELECTRONICA",
                "F001-00000333",
                "Fecha de Emisión: 10/03/2026",
                f"RUC: {RUC_RECEPTOR}",
                "Total: S/ 1.250",
            ]
        )
        res_dec = parse_document([p_dec], target_recipient_ruc=RUC_RECEPTOR)[0]
        self.assertIsNone(res_dec.total)
        self.assertTrue(res_dec.requires_human_review)
        self.assertTrue(any(c.get("check") == "ambiguous_amount_format" for c in res_dec.derived_checks))

    def test_electro_row_items_sum_check_and_zero_qty(self) -> None:
        # Fila de detalle sin columna de precio unitario: valida que no se invente precio y que sume contra cabecera
        p_electro = make_page(
            [
                "COMERCIAL ANDINA S.A.C.",
                f"RUC: {RUC_EMISOR_A}",
                "FACTURA ELECTRONICA",
                "F003-00000444",
                "Fecha de Emisión: 12/03/2026",
                f"Cliente: EMPRESA COMPRADORA S.A. RUC: {RUC_RECEPTOR}",
                "CANT. UNIDAD DESCRIPCION P.UNIT DSCTO TOTAL",
                "2 NIU CAJA CARTON REFORZADO 100.00",
                "0.0 NIU BONIFICACION EMPAQUE 0.00",
                "Op. Gravada: S/ 100.00",
                "IGV 18%: S/ 18.00",
                "Importe Total: S/ 118.00",
            ]
        )
        res_el = parse_document([p_electro], target_recipient_ruc=RUC_RECEPTOR)[0]
        self.assertEqual(res_el.detection_status, DETECTED)
        if res_el.line_items:
            first_item = res_el.line_items[0]
            self.assertEqual(first_item.line_total_normalized, "100.00")
            # Si la línea no trajo unitario explícito ni match columnario completo, no debe inventar unit_price=100.00
            self.assertNotEqual(first_item.unit_price_normalized, "100.00")

        # Si el total de la cabecera difiere de la suma de ítems, se deja alerta en la conciliación secundaria
        p_mismatch_sum = make_page(
            [
                "COMERCIAL ANDINA S.A.C.",
                f"RUC: {RUC_EMISOR_A}",
                "FACTURA ELECTRONICA",
                "F003-00000445",
                "Fecha de Emisión: 12/03/2026",
                f"Cliente: EMPRESA COMPRADORA S.A. RUC: {RUC_RECEPTOR}",
                "CANT. UNIDAD DESCRIPCION P.UNIT TOTAL",
                "2 NIU TORNILLO 10.00 20.00",
                "Op. Gravada: S/ 100.00",
                "IGV 18%: S/ 18.00",
                "Importe Total: S/ 118.00",
            ]
        )
        res_ms = parse_document([p_mismatch_sum], target_recipient_ruc=RUC_RECEPTOR)[0]
        self.assertTrue(any(c.get("check") == "line_item_arithmetic_reconciliation" for c in res_ms.derived_checks))


class TestPadron(unittest.TestCase):
    def test_name_normalization_ruc10_order_and_nombre_comercial(self) -> None:
        # Mismo nombre con abreviatura y "&"
        st1, _ = compare_names_with_padron("COMERCIAL A&B SOCIEDAD ANONIMA CERRADA", "COMERCIAL A Y B S.A.C.", ruc=RUC_EMISOR_A)
        self.assertEqual(st1, "coincide_exacto")

        # Persona natural RUC 10 con orden invertido apellidos <-> nombres
        st2, _ = compare_names_with_padron("JUAN CARLOS PEREZ GOMEZ", "PEREZ GOMEZ JUAN CARLOS", ruc=RUC_NAT_C)
        self.assertIn(st2, {"coincide_exacto", "coincide_con_ruido_ocr"})

        # Posible nombre comercial frente a discrepancia total de razón social.
        # Dos nombres sin ninguna palabra distintiva en común son discrepancia, no nombre comercial.
        st3, _ = compare_names_with_padron("TIENDA LA ECONOMICA", "COMERCIAL DEL SUR DEMO S.A.", ruc=RUC_EMISOR_A)
        self.assertEqual(st3, "discrepancia")
        st3b, _ = compare_names_with_padron("TIENDA LA ECONOMICA", "LA ECONOMICA DEL SUR DEMO S.A.", ruc=RUC_EMISOR_A)
        self.assertEqual(st3b, "posible_nombre_comercial")

        st4, _ = compare_names_with_padron("FERRETERIA LOS ANDES S.R.L.", "TEXTILES PACIFICO S.A.", ruc=RUC_EMISOR_A)
        self.assertEqual(st4, "discrepancia")

    def test_local_padron_adapter_sqlite_present_and_absent(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            db = make_padron_sqlite(
                Path(tmp) / "padron.sqlite3",
                [
                    (RUC_EMISOR_A, "SERVICIOS INDUSTRIALES ANDINOS S.A.C.", "ACTIVO", "HABIDO", "150101"),
                    (RUC_RECEPTOR, "EMPRESA COMPRADORA S.A.", "ACTIVO", "HABIDO", "150101"),
                ],
            )
            adapter = LocalPadronAdapter(db)
            rep = adapter.evaluate_rucs([RUC_EMISOR_A, RUC_EMISOR_B])
            self.assertEqual(rep.entries[RUC_EMISOR_A].status, PADRON_PRESENTE)
            self.assertEqual(rep.entries[RUC_EMISOR_B].status, PADRON_AUSENTE)

            # Verificar carga desde TXT mediante tabla temporal hacia tabla padron
            txt_src = Path(tmp) / "padron_mini.txt"
            txt_src.write_text(
                f"RUC|NOMBRE O RAZON SOCIAL|ESTADO DEL CONTRIBUYENTE|CONDICION DE DOMICILIO|UBIGEO|\n"
                f"{RUC_EMISOR_A}|SERVICIOS INDUSTRIALES ANDINOS S.A.C.|ACTIVO|HABIDO|150101|\n",
                encoding="cp1252",
            )
            st_db = Path(tmp) / "staged.sqlite3"
            st_res = cargar_padron_staging(txt_src, st_db, cutoff_date="2026-03-01")
            self.assertEqual(st_res.unique_rucs, 1)
            self.assertEqual(LocalPadronAdapter(st_db).evaluate_rucs([RUC_EMISOR_A]).entries[RUC_EMISOR_A].status, PADRON_PRESENTE)


class TestProcessorPdfOrganizerExportsAndCLI(unittest.TestCase):
    def test_qr_crosscheck_detects_igv_and_recipient_ruc_mismatch(self) -> None:
        res = ExtractionResult(
            emitter_ruc=RUC_EMISOR_A,
            recipient_ruc=RUC_RECEPTOR,
            document_type="01",
            series="F001",
            correlativo_normalized="00000042",
            issue_date="2026-03-15",
            igv="18.00",
            total="118.00",
            detection_status=DETECTED,
            comprobante_detectado=True,
        )
        qr_diff = f"{RUC_EMISOR_A}|01|F001|00000042|19.00|118.00|15/03/2026|6|{RUC_OTRO}|"
        page = make_page(["FACTURA"], qr_payloads=[qr_diff])
        _crosscheck_sunat_qr(res, [page])
        qr_chk = next(c for c in res.derived_checks if c.get("check") == "sunat_qr_crosscheck")
        self.assertFalse(qr_chk["consistent"])
        self.assertEqual(qr_chk["status"], "difiere")
        joined = " ".join(qr_chk["discrepancies"])
        self.assertIn("igv(", joined)
        self.assertIn("ruc_receptor(", joined)

    def test_end_to_end_scan_expediente_unification_exports_and_cli(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            inp = tmp_path / "entrada"
            out = tmp_path / "salida"
            inp.mkdir()

            padron_db = make_padron_sqlite(
                tmp_path / "padron.sqlite3",
                [
                    (RUC_EMISOR_A, "SERVICIOS INDUSTRIALES ANDINOS S.A.C.", "ACTIVO", "HABIDO", "150101"),
                    (RUC_EMISOR_B, "COMERCIAL DEL SUR S.R.L.", "ACTIVO", "HABIDO", "150101"),
                    (RUC_RECEPTOR, "EMPRESA COMPRADORA S.A.", "ACTIVO", "HABIDO", "150101"),
                    (RUC_OTRO, "TERCERO AJENO S.A.C.", "ACTIVO", "HABIDO", "150101"),
                ],
            )

            # 1) Factura F001-00000042 + página 2 de soporte (Guía de Remisión) en el mismo PDF
            f42_lines = [
                "SERVICIOS INDUSTRIALES ANDINOS S.A.C.",
                f"RUC: {RUC_EMISOR_A}",
                "FACTURA ELECTRONICA",
                "F001-00000042",
                "Fecha de Emisión: 15/03/2026",
                "Fecha de Vencimiento: 15/04/2026",
                "Condición de Pago: Crédito",
                f"Cliente: EMPRESA COMPRADORA S.A. RUC: {RUC_RECEPTOR}",
                "Orden de Compra: OC-7788",
                "Guía de Remisión: T001-00000999",
                "CANT. UNIDAD DESCRIPCION P.UNIT TOTAL",
                "1.00 ZZ SERVICIO DE INGENIERIA 4,000.00 4,000.00",
                "0.0 NIU BONIFICACION EMPAQUE 0.00 0.00",
                "Op. Gravada: S/ 4,000.00",
                "IGV 18%: S/ 720.00",
                "Importe Total: S/ 4,720.00",
            ]
            support_lines = [
                "GUIA DE REMISION REMITENTE",
                "T001-00000999",
                f"Remitente RUC: {RUC_EMISOR_A}",
                f"Destinatario RUC: {RUC_RECEPTOR}",
                "Traslado de equipos para factura F001-00000042",
            ]
            make_text_pdf(inp / "factura_con_soporte.pdf", [f42_lines, support_lines])

            # 2) Nota de crédito FC01-00000005 que modifica F001-00000042 (mismo proveedor)
            nc_lines = [
                "SERVICIOS INDUSTRIALES ANDINOS S.A.C.",
                f"RUC: {RUC_EMISOR_A}",
                "NOTA DE CREDITO ELECTRONICA",
                "FC01-00000005",
                "Fecha de Emisión: 16/03/2026",
                f"Cliente: EMPRESA COMPRADORA S.A. RUC: {RUC_RECEPTOR}",
                "Tipo de Nota: 04 - Descuento global",
                "Documento que modifica: F001-00000042",
                "Op. Gravada: S/ 100.00",
                "IGV 18%: S/ 18.00",
                "Importe Total: S/ 118.00",
            ]
            make_text_pdf(inp / "nc_proveedor_a.pdf", [nc_lines])

            # 3) Duplicado en ZIP de la misma factura (para probar deduplicación a 05_duplicados_secundarios/)
            dup_pdf = make_text_pdf(tmp_path / "dup_inside_zip.pdf", [f42_lines])
            with zipfile.ZipFile(inp / "lote_respaldo.zip", "w") as zf:
                zf.write(dup_pdf, arcname="sub/dup_f42.pdf")

            # 4) Factura emitida a otro RUC (TERCERO_) -> debe quedar en 03_revision_pendiente/
            tercero_lines = [
                "COMERCIAL DEL SUR S.R.L.",
                f"RUC: {RUC_EMISOR_B}",
                "FACTURA ELECTRONICA",
                "F001-00000100",
                "Fecha de Emisión: 18/03/2026",
                f"Cliente: TERCERO AJENO S.A.C. RUC: {RUC_OTRO}",
                "Mercadería diversa",
                "Op. Gravada: S/ 100.00",
                "IGV 18%: S/ 18.00",
                "Importe Total: S/ 118.00",
            ]
            make_text_pdf(inp / "factura_tercero.pdf", [tercero_lines])

            # 5) Proveedor ausente del padrón para probar alerta en hoja revision
            unknown_ruc = make_valid_ruc("2099000066")
            unknown_lines = [
                "PROVEEDOR NUEVO NO PADRON S.A.C.",
                f"RUC: {unknown_ruc}",
                "FACTURA ELECTRONICA",
                "F010-00000001",
                "Fecha de Emisión: 20/03/2026",
                f"Cliente: EMPRESA COMPRADORA S.A. RUC: {RUC_RECEPTOR}",
                "Consultoría rápida",
                "Op. Gravada: S/ 200.00",
                "IGV 18%: S/ 36.00",
                "Importe Total: S/ 236.00",
            ]
            make_text_pdf(inp / "factura_no_en_padron.pdf", [unknown_lines])

            # Ejecutar CLI `scan --exportar-pdf`
            cfg_file = tmp_path / "cfg.json"
            cfg_file.write_text(
                json.dumps(
                    {
                        "unificar_notas_en_proveedor": True,
                        "tolerancia_redondeo": 0.02,
                    }
                ),
                encoding="utf-8",
            )
            rc = cli_main(
                [
                    "scan",
                    "--input",
                    str(inp),
                    "--output",
                    str(out),
                    "--recipient-ruc",
                    RUC_RECEPTOR,
                    "--config",
                    str(cfg_file),
                    "--padron-db",
                    str(padron_db),
                    "--exportar-pdf",
                ]
            )
            self.assertEqual(rc, 0)

            xlsx_path = out / "CDP.xlsx"
            self.assertTrue(xlsx_path.exists())

            # Verificar organización de expediente_pdf/
            exp_root = out / "expediente_pdf"
            self.assertTrue((exp_root / "01_cdp_por_proveedor").exists())
            prov_a_files = [p.name for p in (exp_root / "01_cdp_por_proveedor").rglob("*.pdf")]
            # Factura + Nota unificada + Soporte contiguo del mismo PDF
            self.assertIn(f"{RUC_EMISOR_A}-01-F001-00000042.pdf", prov_a_files)
            self.assertIn(f"{RUC_EMISOR_A}-07-FC01-00000005.pdf", prov_a_files)
            self.assertTrue(any("_SOPORTES_" in name for name in prov_a_files), f"Faltó soporte vinculado: {prov_a_files}")

            # 02_notas_por_proveedor no debe contener notas si unify_notes_by_provider=True
            notas_dir_files = list((exp_root / "02_notas_por_proveedor").rglob("*.pdf")) if (exp_root / "02_notas_por_proveedor").exists() else []
            self.assertEqual(notas_dir_files, [])

            # Tercero ajeno derivado a 03_revision_pendiente/ con prefijo TERCERO_
            rev_files = [p.name for p in (exp_root / "03_revision_pendiente").rglob("*.pdf")]
            self.assertTrue(any(name.startswith("TERCERO_") and "F001-00000100" in name for name in rev_files), f"No derivó TERCERO_: {rev_files}")

            # Duplicado del ZIP en 05_duplicados_secundarios/
            dup_files = [p.name for p in (exp_root / "05_duplicados_secundarios").rglob("*.pdf")]
            self.assertTrue(any("F001-00000042" in name for name in dup_files), f"No archivó duplicado: {dup_files}")

            # Verificar contenido del libro Excel CDP.xlsx
            wb = openpyxl.load_workbook(xlsx_path, data_only=False)
            self.assertEqual(wb.sheetnames, ["cdp_detectados", "notas", "revision", "inventario", "detalle_items", "resumen"])

            ws_cdp = wb["cdp_detectados"]
            cdp_headers = [cell.value for cell in ws_cdp[1]]
            for expected_col in (
                "id_documento",
                "fecha_de_vencimiento",
                "condicion_de_pago",
                "tasa_igv_efectiva",
                "apto_credito_fiscal_preliminar",
                "alerta_bancarizacion",
                "orden_de_compra_citada",
                "guia_de_remision_citada",
                "anexos_mismo_pdf",
                "peso_pdf_mb",
                "supera_10_mb_mpv",
            ):
                self.assertIn(expected_col, cdp_headers)

            # Verificar bancarización en factura > S/ 2,000 y fila de tercero en CDP detectados
            idx_serie = cdp_headers.index("serie")
            idx_corr = cdp_headers.index("correlativo")
            idx_banc = cdp_headers.index("alerta_bancarizacion")
            idx_apto = cdp_headers.index("apto_credito_fiscal_preliminar")
            idx_sop = cdp_headers.index("anexos_mismo_pdf")
            rows_by_key = {
                (row[idx_serie].value, row[idx_corr].value): row
                for row in ws_cdp.iter_rows(min_row=2)
                if row[idx_serie].value is not None
            }
            self.assertIn(("F001", "00000042"), rows_by_key)
            self.assertIn(("F001", "00000100"), rows_by_key)
            row_f42 = rows_by_key[("F001", "00000042")]
            self.assertIn("bancariz", str(row_f42[idx_banc].value).lower())
            self.assertNotEqual(str(row_f42[idx_apto].value), "")
            self.assertIn("SOPORTES", str(row_f42[idx_sop].value))
            row_terc = rows_by_key[("F001", "00000100")]
            self.assertIn("RUC receptor distinto", str(row_terc[idx_apto].value))

            # Verificar hoja detalle_items conserva ID documento, cantidades numéricas (incluyendo 0.0)
            ws_det = wb["detalle_items"]
            det_headers = [cell.value for cell in ws_det[1]]
            self.assertIn("id_documento", det_headers)
            self.assertIn("cantidad", det_headers)
            self.assertGreaterEqual(ws_det.max_row, 2)
            idx_qty = det_headers.index("cantidad")
            qtys = [row[idx_qty].value for row in ws_det.iter_rows(min_row=2)]
            self.assertIn(0, qtys, f"Se esperaba cantidad numérica 0 preservada en detalle_items: {qtys}")

            # Verificar hoja notas con signo negativo en NC 07
            ws_notas = wb["notas"]
            notas_headers = [cell.value for cell in ws_notas[1]]
            idx_ntot = notas_headers.index("total_con_signo")
            ntot_vals = [row[idx_ntot].value for row in ws_notas.iter_rows(min_row=2) if row[idx_ntot].value is not None]
            self.assertTrue(any(float(v) < 0 for v in ntot_vals), f"Se esperaba importe negativo en notas: {ntot_vals}")

            # Verificar que en revision aparezcan PADRON_AUSENTE y tercero/receptor
            ws_rev = wb["revision"]
            rev_text = "\n".join(
                str(cell.value)
                for row in ws_rev.iter_rows(min_row=2)
                for cell in row
                if cell.value is not None
            )
            self.assertIn(unknown_ruc, rev_text)

            # Verificar resumen incluye advertencia de alcance legal/tributario
            ws_res = wb["resumen"]
            res_text = "\n".join(
                str(cell.value)
                for row in ws_res.iter_rows(min_row=1)
                for cell in row
                if cell.value is not None
            )
            self.assertIn("NO consulta Clave SOL", res_text)
            wb.close()

            # Verificar que CLI `export` vuelve a generar el Excel desde SQLite sin error
            rc_export = cli_main(["export", "--output", str(out)])
            self.assertEqual(rc_export, 0)

    def test_large_image_adaptive_compression_and_dedup_priority(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            inp = tmp_path / "entrada"
            out = tmp_path / "salida"
            inp.mkdir()

            # Imagen PNG ruidosa para verificar guardado optimizado en expediente_pdf/
            img_path = inp / "boleta_scan.png"
            img = Image.new("RGB", (1200, 1600), color=(245, 245, 245))
            img.save(img_path, format="PNG")

            scan_text = "\n".join(
                [
                    "COMERCIAL DEL SUR S.R.L.",
                    f"RUC: {RUC_EMISOR_B}",
                    "FACTURA ELECTRONICA",
                    "F005-00000077",
                    "Fecha de Emisión: 10/03/2026",
                    f"Cliente: EMPRESA COMPRADORA S.A. RUC: {RUC_RECEPTOR}",
                    "Servicio logístico",
                    "Op. Gravada: S/ 100.00",
                    "IGV 18%: S/ 18.00",
                    "Importe Total: S/ 118.00",
                ]
            )
            # Además un PDF nativo en ZIP con los mismos datos debe ganar como primario frente al escaneado en disco
            native_in_zip = make_text_pdf(tmp_path / "nat_f77.pdf", [scan_text.splitlines()])
            with zipfile.ZipFile(inp / "nativos.zip", "w") as zf:
                zf.write(native_in_zip, arcname="F005-77.pdf")

            cfg = AppConfig(
                recipient_ruc=RUC_RECEPTOR,
                export_pdf_folders=True,
                unify_notes_by_provider=True,
            )
            processor = FiscalProcessor(out, config=cfg, engine=MotorOCRFicticio(scan_text))
            summary = processor.scan(inp)
            self.assertEqual(summary.detected_count, 2)
            xlsx_path = export_results(processor.store, out, summary=summary)

            wb = openpyxl.load_workbook(xlsx_path, data_only=False)
            ws_cdp = wb["cdp_detectados"]
            headers = [c.value for c in ws_cdp[1]]
            src_idx = headers.index("archivo_de_origen")
            # El primario debe ser el PDF nativo del ZIP (native_pdf_text vence a OCR en disco)
            primaries = [row[src_idx].value for row in ws_cdp.iter_rows(min_row=2) if row[src_idx].value]
            wb.close()
            self.assertEqual(len(primaries), 1)
            self.assertIn("nativos.zip", str(primaries[0]))


if __name__ == "__main__":
    unittest.main()
