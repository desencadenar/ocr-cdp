"""Pruebas de los controles fiscales: detección, cuadre de importes, tasas, cotejo con QR y etiquetas.

Cada prueba arma un documento sintético con el caso que el control debe resolver.
Ningún dato sale de documentos reales: los RUC se construyen con dígito verificador válido.
"""

from __future__ import annotations

import contextlib
from dataclasses import replace
import io
import json
import shutil
import zipfile
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest import mock

import openpyxl
import pymupdf

ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from ocr_fiscal.amounts import amount_mentions, evaluate_igv_rate, reconcile_amounts
from ocr_fiscal.cli import main as cli_main
from ocr_fiscal.config import AppConfig, load_config
from ocr_fiscal.padron import compare_names_with_padron
from ocr_fiscal.exports import _row
from ocr_fiscal.pdf_organizer import _has_blocking_expediente_issue
from ocr_fiscal.documents import QR_DECODER_AVAILABLE, PageRead
from ocr_fiscal.archives import ContainerWalker
from ocr_fiscal.processor import FiscalProcessor, _crosscheck_sunat_qr
from ocr_fiscal.extraction import _normalize_amount_str, parse_document
from ocr_fiscal.models import AMBIGUOUS, DETECTED, IDENTITY_COMPLETE, NOT_RECEIPT, ExtractionResult, LayoutLine


def ruc(prefix10: str) -> str:
    weights = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    check = 11 - sum(int(d) * w for d, w in zip(prefix10, weights)) % 11
    return prefix10 + str({10: 0, 11: 1}.get(check, check))


EM = ruc("2099000033")
RE = ruc("2099000011")
OT = ruc("2099000022")
NAT = ruc("1099000055")


def page(lines: list[str], *, source: str = "native_pdf_text", conf: float | None = None, qr: list[str] | None = None) -> PageRead:
    layout = [
        LayoutLine(
            line_number=i, text=t, block_number=1, bbox=(10.0, i * 15.0, 500.0, i * 15.0 + 12),
            source=source, ocr_confidence=conf,
        )
        for i, t in enumerate(lines, 1)
    ]
    return PageRead(
        page_number=1, text="\n".join(lines), source=source, status="ok",
        quality={"char_count": 100, "line_count": len(lines)}, layout_lines=layout, qr_payloads=list(qr or []),
    )


def parse(lines: list[str], **kw):
    return parse_document([page(lines, **kw)], target_recipient_ruc=RE)


def check(result, name: str) -> dict:
    return next((c for c in result.derived_checks if c.get("check") == name), {})


HEAD = [
    "COMERCIAL ANDINA S.A.C.",
    f"RUC: {EM}",
    "FACTURA ELECTRONICA",
    "F001-00001234",
    "Fecha de Emisión: 15/03/2026",
    f"Señor(es): COMPRADORA S.A. RUC: {RE}",
    "Por concepto de: servicio de mantenimiento",
]


class TestComprobantesNoDescartadosComoSoporte(unittest.TestCase):
    """Un bloque con RUC válido, serie-correlativo e importes no se descarta como soporte sin revisión."""

    def assert_not_silently_dropped(self, results) -> None:
        self.assertTrue(results)
        for r in results:
            self.assertNotEqual(r.detection_status, NOT_RECEIPT, r.decision_comment)
        self.assertTrue(any(r.detection_status in {DETECTED, AMBIGUOUS} and r.requires_human_review for r in results))

    def test_titulo_mal_leido_y_guia_citada(self) -> None:
        self.assert_not_silently_dropped(parse([
            "COMERCIAL ANDINA S.A.C.", f"RUC: {EM}", "FACTVRA ELECTRONICA", "F001-00001234",
            "Fecha de Emisión: 15/03/2026", f"Señor(es): COMPRADORA S.A. RUC: {RE}",
            "GUIA DE REMISION REMITENTE T001-000456",
            "Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00",
        ]))

    def test_recibo_con_estado_de_cuenta_del_suministro(self) -> None:
        self.assert_not_silently_dropped(parse([
            "ELECTRO SUR S.A.A.", f"RUC: {EM}", "RECIBO N° S001-00045678", "ESTADO DE CUENTA DEL SUMINISTRO",
            "Fecha de Emisión: 15/03/2026", f"Cliente: COMPRADORA S.A. RUC: {RE}", "Consumo de energia 100.00",
            "Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00",
        ]))

    def test_factura_fisica_con_aceptacion_de_servicios(self) -> None:
        self.assert_not_silently_dropped(parse([
            "COMERCIAL ANDINA S.A.C.", f"R.U.C. {EM}", "001-0001234", "Fecha: 15/03/2026",
            f"Señor(es): COMPRADORA S.A. RUC: {RE}", "Por la aceptación de servicios de mantenimiento",
            "SUB TOTAL 100.00", "IGV 18.00", "TOTAL 118.00",
        ]))

    def test_sin_titulo_y_glosa_orden_de_compra(self) -> None:
        self.assert_not_silently_dropped(parse([
            "COMERCIAL ANDINA S.A.C.", f"RUC: {EM}", "F001-00001234", "Fecha de Emisión: 15/03/2026",
            f"Señor(es): COMPRADORA S.A. RUC: {RE}", "Atención según ORDEN DE COMPRA 4500012345 del cliente",
            "Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00",
        ]))

    def test_soportes_sin_estructura_de_comprobante_siguen_descartados(self) -> None:
        cases = {
            "guia_de_remision": [
                "TRANSPORTES ANDINOS S.A.C.", f"RUC: {EM}", "GUIA DE REMISION REMITENTE ELECTRONICA", "T001-00000456",
                "Fecha de traslado: 15/03/2026", "Punto de partida: Lima", "1 UND BOMBA HIDRAULICA",
            ],
            "orden_de_compra_pedido": [
                "COMPRADORA S.A.", f"RUC: {RE}", "ORDEN DE COMPRA 4500012345", f"Proveedor: COMERCIAL ANDINA RUC {EM}",
                "Subtotal: 100.00", "IGV: 18.00", "Total: 118.00",
            ],
            "estado_de_cuenta_bancario": [
                "BANCO ANDINO", f"RUC: {EM}", "ESTADO DE CUENTA", "Cuenta 191-1234567-0-12",
                "Saldo anterior 1,000.00", "TOTAL 2,500.00",
            ],
            "hoja_entrada_servicios_hes": [
                "HOJA DE ENTRADA DE SERVICIOS", "N° 1000012345", f"Proveedor RUC {EM}", "Valor neto 100.00",
            ],
        }
        for expected, lines in cases.items():
            with self.subTest(expected):
                results = parse(lines)
                self.assertEqual(len(results), 1)
                self.assertEqual(results[0].detection_status, NOT_RECEIPT)
                self.assertEqual(results[0].non_receipt_type, expected)


def scan(docs: dict[str, list[str]], *, recipient: str = RE) -> dict[str, list[dict]]:
    """Corre `scan` sobre PDF de texto sintéticos y devuelve las hojas del Excel como listas de filas."""
    with tempfile.TemporaryDirectory() as tmp:
        inp, out = Path(tmp) / "entrada", Path(tmp) / "salida"
        inp.mkdir()
        for name, lines in docs.items():
            doc = pymupdf.open()
            pdf_page = doc.new_page(width=595, height=842)
            for i, line in enumerate(lines):
                pdf_page.insert_text((36, 40 + 14 * i), line, fontsize=10, fontname="helv")
            doc.save(inp / f"{name}.pdf")
            doc.close()
        with contextlib.redirect_stdout(io.StringIO()):
            rc = cli_main(["scan", "--input", str(inp), "--output", str(out), "--recipient-ruc", recipient])
        assert rc == 0
        wb = openpyxl.load_workbook(out / "CDP.xlsx", read_only=True)
        sheets = {}
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            sheets[ws.title] = [dict(zip(rows[0], r)) for r in rows[1:]] if rows else []
        wb.close()
        return sheets


def R(**kw) -> ExtractionResult:
    r = ExtractionResult(detection_status=DETECTED, comprobante_detectado=True)
    for k, v in kw.items():
        setattr(r, k, v)
    return r


class TestCuadreSoloConComponentesCompletos(unittest.TestCase):
    """«cuadra» solo cuando se comprobaron los componentes; si no, «no_contradice» y revisión."""

    def assert_not_cuadra(self, chk: dict) -> None:
        self.assertNotIn(chk["status"], {"cuadra", "cuadra_con_redondeo"}, chk)
        self.assertIsNot(chk["consistent"], True, chk)

    def test_gravado_igual_al_total_sin_igv(self) -> None:
        self.assert_not_cuadra(reconcile_amounts(R(document_type="01", gravado="118.00", total="118.00")))
        r = parse(HEAD + ["Op. Gravada: S/ 118.00", "Importe Total: S/ 118.00"])[0]
        self.assert_not_cuadra(check(r, "arithmetic_total_reconciliation"))
        self.assertTrue(r.requires_human_review)
        self.assertIn("amounts", r.secondary_review_fields)

    def test_subtotal_y_total_iguales_sin_componentes(self) -> None:
        chk = reconcile_amounts(R(document_type="01", subtotal="118.00", total="118.00"))
        self.assert_not_cuadra(chk)
        self.assertEqual(chk["status"], "no_contradice")

    def test_c14_nota_de_credito_total_cero_con_la_palabra_cuota(self) -> None:
        self.assert_not_cuadra(reconcile_amounts(R(document_type="07", total="0.00", extracted_text="NOTA DE CREDITO F001-1 PAGO EN UNA CUOTA")))
        self.assert_not_cuadra(reconcile_amounts(R(document_type="07", total="0.00", installments_count="1", extracted_text="credito 1 cuota")))
        r = parse([
            "COMERCIAL ANDINA S.A.C.", f"RUC: {EM}", "NOTA DE CREDITO ELECTRONICA", "FC01-00000012",
            "Fecha de Emisión: 15/03/2026", f"Señor(es): COMPRADORA S.A. RUC: {RE}",
            "Documento que modifica: F001-00001234", "Motivo: Descuento global", "Forma de pago: credito 1 cuota",
            "Importe Total: S/ 0.00",
        ])[0]
        self.assert_not_cuadra(check(r, "arithmetic_total_reconciliation"))
        self.assertTrue(r.requires_human_review)

    def test_nota_tipo_13_explicita_sigue_cuadrando(self) -> None:
        self.assertEqual(reconcile_amounts(R(document_type="07", note_type_code="13", total="0.00"))["status"], "cuadra")
        self.assertEqual(reconcile_amounts(R(document_type="07", total="0.00", extracted_text="Tipo 13: ajuste de fechas"))["status"], "cuadra")

    def test_b9_signos_mixtos_no_se_rescatan_con_valores_absolutos(self) -> None:
        self.assert_not_cuadra(reconcile_amounts(R(document_type="07", gravado="100.00", igv="-18.00", total="118.00")))
        self.assert_not_cuadra(reconcile_amounts(R(document_type="07", gravado="-100.00", igv="18.00", total="118.00")))
        self.assert_not_cuadra(reconcile_amounts(R(document_type="07", gravado="-100.00", igv="18.00", total="-118.00")))
        self.assert_not_cuadra(reconcile_amounts(R(document_type="07", gravado="200.00", igv="-18.00", total="-182.00")))

    def test_convenciones_de_signo_coherentes_siguen_cuadrando(self) -> None:
        # Componentes positivos y total impreso en negativo: convención habitual de una nota de crédito.
        self.assertEqual(reconcile_amounts(R(document_type="07", gravado="100.00", igv="18.00", total="-118.00"))["status"], "cuadra")
        self.assertEqual(reconcile_amounts(R(document_type="07", gravado="-100.00", igv="-18.00", total="-118.00"))["status"], "cuadra")

    def test_cuadres_completos_no_cambian(self) -> None:
        self.assertEqual(reconcile_amounts(R(document_type="01", gravado="100.00", igv="18.00", total="118.00"))["status"], "cuadra")
        self.assertEqual(reconcile_amounts(R(document_type="01", exonerado="100.00", total="100.00"))["status"], "cuadra")
        self.assertEqual(reconcile_amounts(R(document_type="02", subtotal="2000.00", importe_retencion="160.00", total_a_pagar="1840.00"))["status"], "cuadra")
        self.assertEqual(reconcile_amounts(R(document_type="01", gratuitas="150.00", total="0.00"))["status"], "cuadra")


class TestImportesAmbiguos(unittest.TestCase):
    """«10.000» es ambiguo (diez mil o diez) y «57» delante del importe suele ser el «S/» mal leído."""

    def test_miles_con_punto_sin_decimales_va_a_revision(self) -> None:
        r = parse(HEAD + ["Op. Gravada: S/ 10.000", "IGV 18%: S/ 1.800", "Importe Total: S/ 11.800"])[0]
        self.assertNotIn(r.gravado, {"10.00", "10000.00"})
        self.assertNotIn(r.igv, {"1.80", "1800.00"})
        self.assertNotIn(r.total, {"11.80", "11800.00"})
        self.assertTrue(r.requires_human_review)
        ambiguous = [c for c in r.derived_checks if c.get("check") == "ambiguous_amount_format"]
        self.assertEqual({c["field"] for c in ambiguous}, {"Base gravada", "IGV", "Total leído"})
        self.assertTrue(all(c["consistent"] is False and "10.000" in " ".join(c2["observed_value"] for c2 in ambiguous) for c in ambiguous))
        self.assertNotEqual(check(r, "arithmetic_total_reconciliation").get("status"), "cuadra")

    def test_coma_con_un_decimal_es_ambigua(self) -> None:
        r = parse(HEAD + ["Op. Gravada: S/ 100,5", "IGV 18%: S/ 18.09", "Importe Total: S/ 118.59"])[0]
        self.assertNotEqual(r.gravado, "1005.00")
        self.assertTrue(any(c.get("check") == "ambiguous_amount_format" for c in r.derived_checks))

    def test_formatos_inequivocos_no_cambian(self) -> None:
        r = parse(HEAD + ["Op. Gravada: S/ 10,000.00", "IGV 18%: S/ 1,800.00", "Importe Total: S/ 11,800.00"])[0]
        self.assertEqual((r.gravado, r.igv, r.total), ("10000.00", "1800.00", "11800.00"))
        self.assertFalse(any(c.get("check") == "ambiguous_amount_format" for c in r.derived_checks))
        r = parse(HEAD + ["Op. Gravada: S/ 10.000,00", "IGV 18%: S/ 1.800,00", "Importe Total: S/ 11.800,00"])[0]
        self.assertEqual((r.gravado, r.igv, r.total), ("10000.00", "1800.00", "11800.00"))
        self.assertEqual(_normalize_amount_str("1,234"), "1234.00")

    def test_57_delante_del_importe(self) -> None:
        self.assertTrue(list(amount_mentions("TOTAL 57 55.00"))[0][3])
        self.assertFalse(list(amount_mentions("TOTAL 57.00"))[0][3])
        self.assertFalse(list(amount_mentions("TOTAL S/ 57.00"))[0][3])
        r = parse(HEAD + ["Op. Gravada 57 55.00", "IGV 57 9.90", "Importe Total 57 64.90"])[0]
        self.assertNotEqual(r.gravado, "57.00")
        self.assertNotEqual(r.total, "57.00")
        self.assertTrue(r.requires_human_review)


class TestConfianzaOcrEnImportes(unittest.TestCase):
    def test_confianza_baja_en_importes_genera_hallazgo(self) -> None:
        r = parse(HEAD + ["Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00"], source="ocr", conf=0.20)[0]
        fields = {c["field"] for c in r.derived_checks if c.get("check") == "low_ocr_confidence"}
        self.assertTrue({"Base gravada", "IGV", "Total leído"} <= fields, fields)

    def test_confianza_alta_no_genera_hallazgo(self) -> None:
        r = parse(HEAD + ["Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00"], source="ocr", conf=0.95)[0]
        self.assertFalse(any(c.get("check") == "low_ocr_confidence" for c in r.derived_checks))


RHE = [
    "PEREZ LOPEZ JUAN", f"RUC: {NAT}", "RECIBO POR HONORARIOS ELECTRONICO", "Nro: E001-45",
    "Fecha de emisión: 15/03/2026", f"Recibí de: COMPRADORA S.A. RUC: {RE}", "Por concepto de: asesoria contable",
    "Total por Honorarios: 1,000.00", "Retención (8%) IR: 0.00", "Total Neto Recibido: 1,000.00",
]


class TestFaltaDeDatoNoEsConsistente(unittest.TestCase):
    def test_no_evaluable_y_no_aplica_no_devuelven_consistent_true(self) -> None:
        for kw in (
            dict(document_type="01", total="118.00"),
            dict(document_type="01", gravado="100.00", total="118.00"),
            dict(document_type="03", total="50.00"),
            dict(document_type="02", subtotal="1000.00"),                   # no_aplica
        ):
            with self.subTest(kw):
                chk = evaluate_igv_rate(R(**kw))
                self.assertIn(chk["status"], {"no_evaluable", "no_aplica"})
                self.assertIsNone(chk["consistent"])

    def test_factura_sin_igv_legible_va_a_revision_por_tasa(self) -> None:
        chk = evaluate_igv_rate(R(document_type="01", total="118.00"))
        self.assertTrue(chk["needs_review"])
        r = parse(HEAD + ["Importe Total: S/ 118.00"])[0]
        self.assertIn("igv_rate", r.secondary_review_fields)

    def test_no_aplica_y_boleta_sin_desglose_no_generan_revision_por_tasa(self) -> None:
        self.assertFalse(evaluate_igv_rate(R(document_type="02", subtotal="1000.00"))["needs_review"])
        self.assertFalse(evaluate_igv_rate(R(document_type="03", total="50.00"))["needs_review"])
        r = parse(RHE)[0]
        self.assertNotIn("igv_rate", r.secondary_review_fields)

    def test_revision_no_lista_no_aplica_y_si_no_evaluable(self) -> None:
        sheets = scan({"rhe": RHE, "factura": HEAD[:6] + ["Importe Total: S/ 118.00"]})
        kinds = [r["tipo_hallazgo"] for r in sheets["revision"]]
        self.assertNotIn("no_aplica", kinds)
        self.assertIn("no_evaluable", kinds)


def qr_status(result: ExtractionResult, payloads: list[str], **kw) -> dict:
    _crosscheck_sunat_qr(result, [page(["FACTURA"], qr=payloads)], **kw)
    return check(result, "sunat_qr_crosscheck")


FULL = dict(emitter_ruc=EM, document_type="01", series="F001", correlativo_normalized="00000042", igv="18.00", total="118.00", issue_date="2026-03-15")
QR_OK = f"{EM}|01|F001|42|18.00|118.00|2026-03-15|6|{RE}|"


class TestCotejoQr(unittest.TestCase):
    def test_coincide_exige_ruc_serie_y_total(self) -> None:
        chk = qr_status(R(**FULL), [QR_OK])
        self.assertEqual((chk["status"], chk["consistent"]), ("coincide", True))

    def test_solo_tipo_leido_no_es_coincide(self) -> None:
        # Sin serie ni RUC leídos no hay con qué vincular el QR: no se da por coincidente.
        chk = qr_status(R(document_type="01"), [QR_OK])
        self.assertEqual(chk["status"], "qr_no_vinculado")
        self.assertIsNone(chk["consistent"])
        # Vinculado por serie y RUC, pero sin total leído: coincidencia parcial, no «coincide».
        partial = dict(FULL, total=None)
        chk = qr_status(R(**partial), [QR_OK])
        self.assertEqual(chk["status"], "coincide_parcial")
        self.assertIsNone(chk["consistent"])
        self.assertIn("total", chk["reason"])

    def test_qr_de_otro_emisor_no_se_toma_por_defecto(self) -> None:
        r = R(document_type="01", total="118.00")
        chk = qr_status(r, [f"{OT}|01|F999|7|18.00|118.00|2026-03-15|6|{OT}|"])
        self.assertEqual(chk["status"], "qr_no_vinculado")
        self.assertIsNone(chk["consistent"])
        self.assertIn("qr_crosscheck", r.secondary_review_fields)

    def test_con_varios_qr_se_usa_el_que_coincide(self) -> None:
        chk = qr_status(R(**FULL), [f"{OT}|01|F999|7|18.00|50.00|2026-03-15|6|{OT}|", QR_OK])
        self.assertEqual(chk["status"], "coincide")

    def test_coma_decimal_en_qr(self) -> None:
        chk = qr_status(R(**FULL), [f"{EM}|01|F001|42|18,00|118,00|2026-03-15|6|{RE}|"])
        self.assertEqual(chk["status"], "coincide")

    def test_sin_lector_de_qr_se_distingue_de_sin_qr(self) -> None:
        self.assertEqual(qr_status(R(**FULL), [], qr_decoder_available=False)["status"], "qr_no_ejecutado")
        self.assertEqual(qr_status(R(**FULL), [], qr_decoder_available=True)["status"], "sin_qr")

    def test_excel_avisa_que_el_cotejo_no_se_ejecuto(self) -> None:
        if QR_DECODER_AVAILABLE:
            self.skipTest("zxingcpp instalado en este entorno")
        sheets = scan({"factura": HEAD + ["Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00"]})
        self.assertEqual({r["cotejo_qr_sunat"] for r in sheets["cdp_detectados"]}, {"qr_no_ejecutado"})
        resumen = {r["indicador"]: r["valor"] for r in sheets["resumen"]}
        self.assertIn("zxingcpp", str(resumen.get("Cotejo con QR SUNAT")))


def rate(**kw) -> dict:
    return evaluate_igv_rate(R(**kw))


class TestControlDeTasa(unittest.TestCase):
    def test_a8_tolerancia_absoluta(self) -> None:
        self.assertEqual(rate(document_type="01", gravado="1000000.00", igv="178001.00", issue_date="2026-03-01")["status"], "tasa_igv_atipica")
        self.assertEqual(rate(document_type="01", gravado="100000.00", igv="17800.00", issue_date="2026-03-01")["status"], "tasa_igv_atipica")
        # Redondeo por línea de unos céntimos sigue siendo 18 %.
        self.assertEqual(rate(document_type="01", gravado="100000.00", igv="18000.07", issue_date="2026-03-01")["status"], "tasa_valida")

    def test_cuatro_por_ciento_solo_con_ivap(self) -> None:
        chk = rate(document_type="01", gravado="1000.00", igv="40.00", issue_date="2026-03-01")
        self.assertEqual(chk["status"], "tasa_igv_atipica")
        self.assertIn("IVAP", chk["reason"])
        chk = rate(document_type="01", gravado="1000.00", igv="40.00", issue_date="2026-03-01", extracted_text="ARROZ PILADO - IVAP 4%")
        self.assertEqual((chk["status"], chk["matched_rate"]), ("tasa_valida", 4.0))

    def test_a4_tasas_especiales_se_marcan_para_verificar(self) -> None:
        for kw in (
            dict(document_type="01", gravado="1000.00", igv="100.00", issue_date="2025-03-01"),
            dict(document_type="01", gravado="1000.00", igv="105.00", issue_date="2026-03-01"),
            dict(document_type="01", gravado="1000.00", igv="150.00", issue_date="2027-03-01"),
        ):
            with self.subTest(kw):
                chk = rate(**kw)
                self.assertEqual(chk["status"], "tasa_especial_verificar")
                self.assertIsNone(chk["consistent"])
                self.assertTrue(chk["needs_review"])
                self.assertIn("restaurantes", chk["reason"])

    def test_tasa_especial_fuera_de_su_anio(self) -> None:
        self.assertEqual(rate(document_type="01", gravado="1000.00", igv="100.00", issue_date="2026-03-01")["status"], "tasa_igv_atipica")

    def test_sin_fecha_no_se_acepta_ninguna_tasa_especial_como_valida(self) -> None:
        chk = rate(document_type="01", gravado="1000.00", igv="150.00")
        self.assertEqual(chk["status"], "tasa_especial_verificar")
        self.assertIn("fecha", chk["reason"])
        self.assertEqual(rate(document_type="01", gravado="1000.00", igv="180.00")["status"], "tasa_valida")

    def test_notas_solo_admiten_la_tasa_del_anio_anterior(self) -> None:
        self.assertEqual(rate(document_type="07", gravado="1000.00", igv="100.00", issue_date="2026-01-15")["status"], "tasa_especial_verificar")
        self.assertEqual(rate(document_type="07", gravado="1000.00", igv="100.00", issue_date="2028-03-01")["status"], "tasa_igv_atipica")
        self.assertEqual(rate(document_type="07", gravado="1000.00", igv="150.00", issue_date="2028-03-01")["status"], "tasa_especial_verificar")
        self.assertEqual(rate(document_type="07", gravado="1000.00", igv="150.00", issue_date="2029-03-01")["status"], "tasa_igv_atipica")

    def test_tasa_especial_va_a_revision(self) -> None:
        r = parse(HEAD + ["Op. Gravada: S/ 1,000.00", "IGV: S/ 105.00", "Importe Total: S/ 1,105.00"])[0]
        self.assertEqual(check(r, "igv_rate_check")["status"], "tasa_especial_verificar")
        self.assertIn("igv_rate", r.secondary_review_fields)
        self.assertTrue(r.requires_human_review)


def apto(doc_type: str, **doc) -> str:
    base = {
        "detection_status": DETECTED, "comprobante_detectado": True,
        "fields_normalized": {"document_type": doc_type}, "recipient_role_status": "match", "recipient_ruc_match": True,
    }
    base.update(doc)
    return _row(base)["apto_credito_fiscal_preliminar"]


class TestCreditoFiscalPorTipo(unittest.TestCase):
    def test_ticket_12_no_recibe_un_no_categorico(self) -> None:
        label = apto("12")
        self.assertFalse(label.startswith("No"), label)
        self.assertIn("RUC del adquirente", label)
        self.assertIn("IGV", label)

    def test_ticket_12_con_receptor_distinto_sigue_siendo_no(self) -> None:
        self.assertTrue(apto("12", recipient_role_status="mismatch").startswith("No: RUC receptor distinto"))

    def test_boleto_terrestre_16_marca_duda_normativa(self) -> None:
        self.assertIn("duda normativa", apto("16"))

    def test_boleta_03_sin_cambios(self) -> None:
        self.assertTrue(apto("03").startswith("No: Boleta"))


class TestLeyendasYTelefonosNoCompitenComoSerie(unittest.TestCase):
    """«R.S. 037-2002/SUNAT» y «(511) 634 - 3400» no son serie-correlativo."""

    FOOTERS = {
        "resolucion_con_barra": "SOMOS AGENTE DE RETENCION, de acuerdo a la R.S. 037-2002/SUNAT",
        "resolucion_sin_barra": "Agente de retención según Resolución de Superintendencia N° 037-2002",
        "rs_sin_puntos": "Designados agentes de percepción RS 128-2002",
        "ley": "Operación sujeta a la Ley 300-2020",
        "decreto": "Exonerado según D.S. 055-1999-EF",
        "telefono_con_area": "Central telefónica (511) 634 - 3400",
        "telefono_rotulado": "Telf.: 634-3400 Anexo 120",
        "telefono_tlf": "TLF 01 634-3400",
    }

    def test_cada_patron_no_genera_conflicto(self) -> None:
        for name, footer in self.FOOTERS.items():
            for header in (HEAD, [line.replace("F001-00001234", "001-0001234") for line in HEAD]):
                with self.subTest(name, serie=header[3]):
                    r = parse(header[:3] + [footer] + header[3:] + ["Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00", footer])[0]
                    self.assertEqual(r.identity_status, IDENTITY_COMPLETE, (r.identity_reason, r.candidates))
                    self.assertEqual(r.series, header[3].split("-")[0])

    def test_serie_numerica_legitima_sigue_leyendose(self) -> None:
        r = parse(["COMERCIAL ANDINA S.A.C.", f"R.U.C. {EM}", "FACTURA", "001-0012345", "Fecha: 15/03/2026",
                   f"Señor(es): COMPRADORA S.A. R.U.C. {RE}", "SUB TOTAL 100.00", "I.G.V. 18.00", "TOTAL S/ 118.00"])[0]
        self.assertEqual(r.series, "001")


CLEAN = HEAD + ["Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00"]


class TestUnaSolaFuenteDeVerdadParaRevision(unittest.TestCase):
    def test_requiere_revision_humana_coincide_con_la_hoja_revision(self) -> None:
        # Cada documento con su propio correlativo: si comparten clave, la deduplicación ya los marca.
        sin_receptor = [line.replace("00001234", "00000002") for line in CLEAN if "Señor(es)" not in line]
        sheets = scan({"limpia": CLEAN, "sin_receptor": sin_receptor})
        ids = {r["id_documento"] for r in sheets["revision"]}
        self.assertTrue(ids)
        for row in sheets["cdp_detectados"]:
            with self.subTest(row["id_documento"]):
                # El Excel guarda la marca como texto («True»/«False»).
                self.assertEqual(str(row["requiere_revision_humana"]) == "True", row["id_documento"] in ids)


VENTA = ["COMPRADORA S.A.", f"RUC: {RE}", "FACTURA ELECTRONICA", "F001-00000077", "Fecha de Emisión: 15/03/2026",
         f"Señor(es): CLIENTE X S.A. RUC: {OT}", "Por concepto de: venta de servicios",
         "Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00"]


class TestVentaPropia(unittest.TestCase):
    def test_venta_propia_se_etiqueta_como_tal_y_no_se_cuenta_dos_veces(self) -> None:
        sheets = scan({"venta": VENTA, "compra": CLEAN})
        venta = next(r for r in sheets["cdp_detectados"] if r["serie"] == "F001" and r["correlativo"] == "00000077")
        self.assertIn("posible venta propia", venta["apto_credito_fiscal_preliminar"].lower())
        kinds = [r["tipo_hallazgo"] for r in sheets["revision"] if r["id_documento"] == venta["id_documento"]]
        self.assertIn("posible_venta_propia_en_compras", kinds)
        self.assertNotIn("ruc_receptor_distinto", kinds)
        resumen = {r["indicador"]: r["valor"] for r in sheets["resumen"]}
        self.assertEqual(resumen["CDP con RUC receptor distinto del objetivo"], 0)

    def test_carpeta_del_expediente_prioriza_venta_propia(self) -> None:
        item = SimpleNamespace(recipient_role_status="mismatch", fields={"emitter_ruc": RE}, derived_checks=[])
        self.assertEqual(_has_blocking_expediente_issue(item, RE), (True, "VENTA_PROPIA_"))


OK_CHECKS = [
    {"check": "arithmetic_total_reconciliation", "method": "observed_components", "status": "cuadra", "consistent": True},
    {"check": "igv_rate_check", "status": "tasa_valida", "consistent": True},
]


class TestLenguajeQueSuenaAConclusion(unittest.TestCase):
    def test_sin_observaciones_dice_pendiente_validez(self) -> None:
        label = apto("01", derived_checks=OK_CHECKS)
        self.assertEqual(label, "Revisar: RUC receptor coincide; pendiente validez CPE y SPOT")
        self.assertNotIn("RUC OK", label)

    def test_con_observaciones_no_emite_pendiente_validez(self) -> None:
        cases = {
            "cuadre": [dict(OK_CHECKS[0], status="no_contradice", consistent=None), OK_CHECKS[1]],
            "tasa IGV": [OK_CHECKS[0], {"check": "igv_rate_check", "status": "tasa_especial_verificar", "consistent": None, "needs_review": True}],
            "lectura OCR": [*OK_CHECKS, {"check": "low_ocr_confidence", "consistent": False, "field": "Total leído"}],
            "padrón": [*OK_CHECKS, {"check": "padron_name_crosscheck", "consistent": False, "emitter_status": "discrepancia"}],
        }
        for problem, checks in cases.items():
            with self.subTest(problem):
                label = apto("01", derived_checks=checks)
                self.assertNotIn("pendiente validez", label)
                self.assertIn(problem, label)
        self.assertIn("padrón", apto("01", derived_checks=OK_CHECKS, secondary_review_fields=["emitter_ruc_padron"]))


def column_page(rows: list[tuple[str, str | None]]) -> PageRead:
    """Página con rótulos en x=28 y valores en x=85 a la misma altura, en el orden en que PyMuPDF los entrega."""
    layout, n = [], 0
    for i, (label, value) in enumerate(rows):
        y = 30.0 + 10 * i
        for x0, x1, text in ((28.0, 60.0, label), (85.0, 300.0, value)):
            if text is None:
                continue
            n += 1
            layout.append(LayoutLine(line_number=n, text=text, block_number=1, bbox=(x0, y, x1, y + 8), source="native_pdf_text"))
    return PageRead(page_number=1, text="\n".join(l.text for l in layout), source="native_pdf_text", status="ok",
                    quality={"char_count": 100, "line_count": len(layout)}, layout_lines=layout, qr_payloads=[])


class TestAdquirenteEnColumnas(unittest.TestCase):
    def build(self, client_label: str) -> PageRead:
        return column_page([
            ("COMERCIAL ANDINA S.A.C.", None), (f"RUC: {EM}", None), ("FACTURA ELECTRONICA", None), ("F001-00001234", None),
            ("Fecha", ": 15/03/2026"), (client_label, ": COMPRADORA S.A."), ("R.U.C.", f": {RE}"), ("Dirección", ": Av. Lima 123"),
            ("Por concepto de: servicio de mantenimiento", None),
            ("Op. Gravada: S/ 100.00", None), ("IGV 18%: S/ 18.00", None), ("Importe Total: S/ 118.00", None),
        ])

    def test_ruc_objetivo_bajo_rotulo_ruc_en_otra_columna(self) -> None:
        for label in ("Señore(s)", "Señor(es)", "Señores", "SEÑORE(S)", "Cliente"):
            with self.subTest(label):
                r = parse_document([self.build(label)], target_recipient_ruc=RE)[0]
                self.assertEqual(r.recipient_ruc, RE, r.observations)
                self.assertEqual(r.recipient_role_status, "verified_recipient")
                self.assertEqual(r.emitter_ruc, EM)


def rhe(amount: str, *, concept: str = "asesoria contable", extra: list[str] | None = None, cur: str = "") -> list[str]:
    return [
        "PEREZ LOPEZ JUAN", f"RUC: {NAT}", "RECIBO POR HONORARIOS ELECTRONICO", "Nro: E001-46", "Fecha de emisión: 15/03/2026",
        f"Recibí de: COMPRADORA S.A. RUC: {RE}", f"Por concepto de: {concept}", *(extra or []),
        f"Total por Honorarios: {cur}{amount}", f"Retención (8%) IR: {cur}0.00", f"Total Neto Recibido: {cur}{amount}",
    ]


def has_rhe_alert(lines: list[str]) -> bool:
    return any(c.get("check") == "rhe_retencion_check" for c in parse(lines)[0].derived_checks)


class TestRetencionHonorarios(unittest.TestCase):
    def test_suspension_en_la_glosa_no_apaga_el_aviso(self) -> None:
        self.assertTrue(has_rhe_alert(rhe("3,000.00", concept="asesoria por suspension de obra")))

    def test_constancia_de_suspension_si_apaga_el_aviso(self) -> None:
        for extra in (["Constancia de suspensión de retenciones N° 123456"], ["Suspensión de retención de 4ta categoría vigente"],
                      ["Formulario Virtual N° 1609 presentado"]):
            with self.subTest(extra):
                self.assertFalse(has_rhe_alert(rhe("3,000.00", extra=extra)))

    def test_umbral_en_soles(self) -> None:
        self.assertFalse(has_rhe_alert(rhe("1,400.00")))
        self.assertTrue(has_rhe_alert(rhe("1,600.00")))

    def test_dolares_sin_tipo_de_cambio_van_a_revision(self) -> None:
        r = parse(rhe("600.00", cur="US$ "))[0]
        chk = check(r, "rhe_retencion_check")
        self.assertTrue(chk, r.currency)
        self.assertIn("tipo de cambio", chk["reason"])
        self.assertTrue(r.requires_human_review)

    def test_dolares_con_tipo_de_cambio_se_convierten(self) -> None:
        self.assertFalse(has_rhe_alert(rhe("300.00", cur="US$ ", extra=["Tipo de cambio: 3.750"])))
        self.assertTrue(has_rhe_alert(rhe("600.00", cur="US$ ", extra=["Tipo de cambio: 3.750"])))


class TestNombresContraPadron(unittest.TestCase):
    def test_palabra_o_numero_distinto_no_es_ruido_ocr(self) -> None:
        for a, b in (
            ("INVERSIONES SAN MARTIN S.A.C.", "INVERSIONES SAN MARCOS S.A.C."),
            ("CONSTRUCTORA NORTE 1 S.A.C.", "CONSTRUCTORA NORTE 2 S.A.C."),
            ("PEREZ LOPEZ JUAN CARLOS", "PEREZ LOPEZ JUAN ALBERTO"),
            ("TRANSPORTES RODRIGUEZ E.I.R.L.", "TRANSPORTES DOMINGUEZ E.I.R.L."),
            ("SERVICIOS GENERALES LIMA S.A.C.", "SERVICIOS GENERALES LIMA NORTE S.A.C."),
        ):
            with self.subTest(a=a, b=b):
                self.assertEqual(compare_names_with_padron(a, b, ruc=EM)[0], "discrepancia")

    def test_nombres_sin_relacion_no_son_nombre_comercial(self) -> None:
        self.assertEqual(compare_names_with_padron("FERRETERIA EL TORNILLO", "GRUPO INDUSTRIAL LIMA S.A.", ruc=EM)[0], "discrepancia")

    def test_ruido_ocr_real_y_nombre_comercial_siguen_reconocidos(self) -> None:
        self.assertEqual(compare_names_with_padron("INVERSI0NES SAN MARTIN S.A.C.", "INVERSIONES SAN MARTIN S.A.C.", ruc=EM)[0], "coincide_con_ruido_ocr")
        self.assertEqual(compare_names_with_padron("INVERSIONES SANMARTIN S.A.C.", "INVERSIONES SAN MARTIN S.A.C.", ruc=EM)[0], "coincide_con_ruido_ocr")
        self.assertEqual(compare_names_with_padron("SERVICIOS GENERALES LIMA NOR", "SERVICIOS GENERALES LIMA NORTE S.A.C.", ruc=EM)[0], "coincide_exacto")
        self.assertEqual(compare_names_with_padron("TIENDA LA ECONOMICA", "LA ECONOMICA DEL SUR S.A.C.", ruc=EM)[0], "posible_nombre_comercial")
        self.assertEqual(compare_names_with_padron("JUAN CARLOS PEREZ GOMEZ", "PEREZ GOMEZ JUAN CARLOS", ruc=NAT)[0], "coincide_exacto")

    def test_padron_sin_antiguedad_maxima_por_defecto(self) -> None:
        # El padrón no tiene límite de antigüedad.
        self.assertIsNone(load_config(None).padron_max_age_days)
        self.assertIsNone(AppConfig().padron_max_age_days)


class TestOrdenDeCompraNoParteLaFactura(unittest.TestCase):
    def test_linea_de_orden_de_compra_dentro_de_la_factura(self) -> None:
        results = parse(HEAD + ["ORDEN DE COMPRA 4500012345", "Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00"])
        self.assertEqual(len(results), 1)
        self.assertEqual((results[0].gravado, results[0].igv, results[0].total), ("100.00", "18.00", "118.00"))

    def test_soporte_completo_despues_de_la_factura_sigue_separandose(self) -> None:
        results = parse(CLEAN + [
            "ORDEN DE COMPRA 4500012345", f"Proveedor: COMERCIAL ANDINA RUC {EM}", "Fecha de entrega: 20/03/2026",
            "Condiciones generales de la orden", "Total orden: 500.00",
        ])
        self.assertEqual([r.detection_status for r in results], [DETECTED, NOT_RECEIPT])
        self.assertEqual(results[0].total, "118.00")


class TestToleranciaUniforme(unittest.TestCase):
    def test_b3_un_centimo_entre_subtotal_y_bases(self) -> None:
        for kw in (dict(subtotal="100.01"), dict(valor_venta="100.01"), dict(subtotal="99.99", valor_venta="99.99")):
            with self.subTest(kw):
                chk = reconcile_amounts(R(document_type="01", gravado="100.00", igv="18.00", total="118.00", **kw))
                self.assertEqual(chk["status"], "cuadra_con_redondeo", chk)
                self.assertTrue(chk["consistent"])
                self.assertIn("redondeo", chk["reason"])

    def test_diferencia_mayor_a_la_tolerancia_sigue_sin_cuadrar(self) -> None:
        chk = reconcile_amounts(R(document_type="01", subtotal="100.05", gravado="100.00", igv="18.00", total="118.00"))
        self.assertIsNot(chk["consistent"], True)


class TestEstadoDeCorrida(unittest.TestCase):
    def run_scan(self, tmp: Path, settings: dict | None) -> Path:
        inp, out = tmp / "entrada", tmp / "salida"
        inp.mkdir()
        doc = pymupdf.open()
        pdf_page = doc.new_page()
        for i, line in enumerate(CLEAN):
            pdf_page.insert_text((36, 40 + 14 * i), line, fontsize=10, fontname="helv")
        doc.save(inp / "f.pdf")
        doc.close()
        args = ["scan", "--input", str(inp), "--output", str(out), "--recipient-ruc", RE]
        if settings is not None:
            (tmp / "cfg.json").write_text(json.dumps(settings), encoding="utf-8")
            args += ["--config", str(tmp / "cfg.json")]
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(cli_main(args), 0)
        return out

    def test_por_defecto_se_conserva_el_estado(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = self.run_scan(Path(tmp), None)
            self.assertTrue((out / ".ocr_fiscal" / "state.sqlite3").is_file())

    def test_opcion_para_no_conservar_el_estado(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = self.run_scan(Path(tmp), {"conservar_estado_corrida": False})
            self.assertTrue((out / "CDP.xlsx").is_file())
            self.assertFalse((out / ".ocr_fiscal").exists())

    def test_la_opcion_se_valida(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            cfg = Path(tmp) / "cfg.json"
            cfg.write_text(json.dumps({"conservar_estado_corrida": "no"}), encoding="utf-8")
            with self.assertRaises(ValueError):
                load_config(cfg)


class TestZipNoSeRetieneEnMemoria(unittest.TestCase):
    def make_pdf_bytes(self, lines: list[str]) -> bytes:
        doc = pymupdf.open()
        pdf_page = doc.new_page()
        for i, line in enumerate(lines):
            pdf_page.insert_text((36, 40 + 14 * i), line, fontsize=10, fontname="helv")
        data = doc.tobytes()
        doc.close()
        return data

    def test_miembros_de_zip_van_a_un_temporal_que_se_borra_al_terminar(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            inp, out = Path(tmp) / "entrada", Path(tmp) / "salida"
            inp.mkdir()
            pdf_bytes = self.make_pdf_bytes(CLEAN)
            with zipfile.ZipFile(inp / "lote.zip", "w") as zf:
                zf.writestr("a/f1.pdf", pdf_bytes)
                zf.writestr("a/f2.pdf", self.make_pdf_bytes([line.replace("00001234", "00000099") for line in CLEAN]))
            walk = ContainerWalker(AppConfig(), output_root=out).walk(inp)
            members = [u for u in walk.work_units if u.source_ref.member_path]
            self.assertEqual(len(members), 2)
            self.assertTrue(all(u.content_bytes is None and u.spool_path and u.spool_path.is_file() for u in members))
            self.assertEqual(members[0].read_bytes(), pdf_bytes)
            shutil.rmtree(walk.spool_dir)

            created: list[str] = []
            real_mkdtemp = tempfile.mkdtemp

            def recording_mkdtemp(*args, **kwargs):
                created.append(real_mkdtemp(*args, **kwargs))
                return created[-1]

            processor = FiscalProcessor(out, replace(AppConfig(), recipient_ruc=RE))
            with mock.patch("ocr_fiscal.archives.tempfile.mkdtemp", recording_mkdtemp):
                summary = processor.scan(inp)
            self.assertEqual(summary.detected_count, 2)
            self.assertEqual(len(created), 1)
            self.assertFalse(Path(created[0]).exists())


if __name__ == "__main__":
    unittest.main()
