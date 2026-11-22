"""Pruebas de precisión de lectura: tasa especial por CIIU, padrón de consulta_ruc_masivo y rótulo del cliente.

Todos los datos son sintéticos. Los RUC se construyen con dígito verificador válido y los archivos de consulta
imitan el formato de salida de la consulta múltiple de SUNAT (separador «|», columna NumeroRuc).
"""

from __future__ import annotations

import contextlib
import io
import os
import sqlite3
import sys
import tempfile
import time
import unittest
from pathlib import Path

import openpyxl
import pymupdf

ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from ocr_fiscal.amounts import evaluate_igv_rate
from ocr_fiscal.ciiu import EmitterActivity, load_consulta_ruc, refine_special_rate_check
from ocr_fiscal.cli import main as cli_main
from ocr_fiscal.config import AppConfig, load_config
from ocr_fiscal.padron import LocalPadronAdapter
from ocr_fiscal.documents import PageRead
from ocr_fiscal.extraction import fuzzy_recipient_label, parse_document
from ocr_fiscal.models import DETECTED, ExtractionResult, LayoutLine


def ruc(prefix10: str) -> str:
    weights = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    check = 11 - sum(int(d) * w for d, w in zip(prefix10, weights)) % 11
    return prefix10 + str({10: 0, 11: 1}.get(check, check))


EM = ruc("2099000033")
RE = ruc("2099000011")
OT = ruc("2099000022")

HEADER = "NumeroRuc|RazonSocial|EstadoContribuyente|CondicionDomicilio|FechaInicioActividades|PrincipalCIIU|SecundarioCIIU1|SecundarioCIIU2"


def consulta_txt(tmp: Path, rows: list[str], *, header: str = HEADER, name: str = "rucs_output.txt") -> Path:
    path = tmp / name
    path.write_text("\n".join([header, *rows]) + "\n", encoding="utf-8")
    return path


def R(**kw) -> ExtractionResult:
    r = ExtractionResult(detection_status=DETECTED, comprobante_detectado=True)
    for k, v in kw.items():
        setattr(r, k, v)
    return r


def activity(principal: str | None = "5610", secondary: tuple[str, ...] = (), **kw) -> EmitterActivity:
    base = dict(ruc=EM, principal=principal, secondary=secondary, estado="ACTIVO", inicio_actividades=None,
                source="rucs_output.txt", query_date="2026-11-18")
    base.update(kw)
    return EmitterActivity(**base)


def special(year: str, igv: str, *, doc_type: str = "01", act: EmitterActivity | None = None, date: str | None = None) -> dict:
    issue = date or f"{year}-03-01"
    chk = evaluate_igv_rate(R(document_type=doc_type, gravado="1000.00", igv=igv, issue_date=issue))
    return refine_special_rate_check(chk, act, issue_date=issue, document_type=doc_type)


class TestTablaDeTasasPorFecha(unittest.TestCase):
    def test_tasa_especial_empieza_en_septiembre_de_2022(self) -> None:
        chk = evaluate_igv_rate(R(document_type="01", gravado="1000.00", igv="100.00", issue_date="2022-08-31"))
        self.assertEqual(chk["status"], "tasa_igv_atipica")
        chk = evaluate_igv_rate(R(document_type="01", gravado="1000.00", igv="100.00", issue_date="2022-09-01"))
        self.assertEqual(chk["status"], "tasa_especial_verificar")

    def test_regimen_termina_en_2027(self) -> None:
        chk = evaluate_igv_rate(R(document_type="01", gravado="1000.00", igv="150.00", issue_date="2028-01-02"))
        self.assertEqual(chk["status"], "tasa_igv_atipica")


class TestLecturaDeLaConsultaRuc(unittest.TestCase):
    def test_txt_con_separador_barra(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = consulta_txt(Path(tmp), [f"{EM}|POLLERIA DEMO S.A.C.|ACTIVO|HABIDO|15/02/2015|5610|4711|"])
            report = load_consulta_ruc(path)
        act = report.activities[EM]
        self.assertEqual((act.principal, act.secondary, act.estado), ("5610", ("4711",), "ACTIVO"))
        self.assertEqual(act.inicio_actividades, "2015-02-15")
        self.assertEqual(act.source, "rucs_output.txt")
        self.assertRegex(act.query_date or "", r"^\d{4}-\d{2}-\d{2}$")

    def test_xlsx_con_rotulos_largos_y_codigo_con_descripcion(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "rucs_output.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "RUCs_SUNAT"
            ws.append(["Número de RUC", "Estado del Contribuyente", "Fecha de Inicio de Actividades",
                       "Actividad Económica Principal", "Actividad Económica Secundaria 1"])
            ws.append([EM, "ACTIVO", "2015-02-15", "5510 - ACTIVIDADES DE ALOJAMIENTO PARA ESTANCIAS CORTAS", "5610 - RESTAURANTES"])
            wb.save(path)
            report = load_consulta_ruc(path)
        act = report.activities[EM]
        self.assertEqual((act.principal, act.secondary), ("5510", ("5610",)))

    def test_archivo_sin_columna_de_ciiu_se_rechaza_con_motivo(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = consulta_txt(Path(tmp), [f"{EM}|DEMO|ACTIVO"], header="NumeroRuc|RazonSocial|EstadoContribuyente")
            with self.assertRaisesRegex(ValueError, "CIIU"):
                load_consulta_ruc(path)

    def test_ruc_mal_formado_se_ignora(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = consulta_txt(Path(tmp), [f"{EM}|A|ACTIVO|HABIDO||5610||", "12345|B|ACTIVO|HABIDO||5610||"])
            report = load_consulta_ruc(path)
        self.assertEqual(list(report.activities), [EM])
        self.assertEqual(report.rejected_rows, 1)


class TestReglaPorCiiuYAnio(unittest.TestCase):
    def test_ciiu_incluido_y_tasa_del_anio_es_valida(self) -> None:
        for year, igv in (("2025", "100.00"), ("2026", "105.00"), ("2027", "150.00")):
            for code in ("5610", "5510"):
                with self.subTest(year=year, code=code):
                    chk = special(year, igv, act=activity(code))
                    self.assertEqual(chk["status"], "tasa_especial_valida")
                    self.assertIs(chk["consistent"], True)
                    self.assertFalse(chk["needs_review"])
                    self.assertIn(code, chk["reason"])

    def test_ciiu_incluido_y_tasa_de_otro_anio_es_atipica(self) -> None:
        self.assertEqual(special("2026", "100.00", act=activity("5610"))["status"], "tasa_igv_atipica")
        self.assertEqual(special("2028", "150.00", act=activity("5610"))["status"], "tasa_igv_atipica")

    def test_ciiu_no_incluido_es_atipica_con_motivo_de_giro(self) -> None:
        chk = special("2026", "105.00", act=activity("4711"))
        self.assertEqual(chk["status"], "tasa_igv_atipica")
        self.assertIs(chk["consistent"], False)
        self.assertTrue(chk["needs_review"])
        self.assertIn("el giro del emisor no es restaurante, hotel ni alojamiento", chk["reason"])

    def test_catering_y_concesionario_excluidos(self) -> None:
        for code in ("5621", "5629"):
            with self.subTest(code=code):
                chk = special("2026", "105.00", act=activity(code))
                self.assertEqual(chk["status"], "tasa_igv_atipica")
                self.assertIn("excluid", chk["reason"])

    def test_sin_dato_de_ciiu_queda_para_verificar(self) -> None:
        chk = special("2026", "105.00", act=None)
        self.assertEqual(chk["status"], "tasa_especial_verificar")
        self.assertTrue(chk["needs_review"])
        self.assertIn("no figura en la consulta RUC", chk["reason"])

    def test_solo_actividad_secundaria_incluida_queda_para_verificar(self) -> None:
        chk = special("2026", "105.00", act=activity("4711", ("5610",)))
        self.assertEqual(chk["status"], "tasa_especial_verificar")
        self.assertIn("secundaria", chk["reason"])

    def test_ciiu_revision_3_queda_para_verificar(self) -> None:
        chk = special("2026", "105.00", act=activity("55205"))
        self.assertEqual(chk["status"], "tasa_especial_verificar")
        self.assertIn("Rev. 3", chk["reason"])

    def test_inicio_de_actividades_posterior_o_estado_no_activo_van_a_revision(self) -> None:
        chk = special("2026", "105.00", act=activity("5610", inicio_actividades="2026-06-01"))
        self.assertEqual(chk["status"], "tasa_especial_verificar")
        self.assertIn("inicio de actividades", chk["reason"])
        chk = special("2026", "105.00", act=activity("5610", estado="BAJA DE OFICIO"))
        self.assertEqual(chk["status"], "tasa_especial_verificar")
        self.assertIn("BAJA DE OFICIO", chk["reason"])

    def test_notas_con_la_tasa_del_anio_anterior(self) -> None:
        chk = special("2026", "100.00", doc_type="07", act=activity("5610"), date="2026-01-15")
        self.assertEqual(chk["status"], "tasa_especial_verificar")
        self.assertIn("año anterior", chk["reason"])
        chk = special("2026", "105.00", doc_type="07", act=activity("5610"), date="2026-01-15")
        self.assertEqual(chk["status"], "tasa_especial_valida")
        chk = special("2026", "100.00", doc_type="07", act=activity("4711"), date="2026-01-15")
        self.assertEqual(chk["status"], "tasa_igv_atipica")

    def test_tasa_general_no_cambia(self) -> None:
        chk = special("2026", "180.00", act=activity("4711"))
        self.assertEqual(chk["status"], "tasa_valida")


def scan(docs: dict[str, list[str]], *extra: str, recipient: str = RE) -> dict[str, list[dict]]:
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
        args = [a.replace("{tmp}", tmp) for a in extra]
        with contextlib.redirect_stdout(io.StringIO()):
            rc = cli_main(["scan", "--input", str(inp), "--output", str(out), "--recipient-ruc", recipient, *args])
        assert rc == 0
        wb = openpyxl.load_workbook(out / "CDP.xlsx", read_only=True)
        sheets = {}
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            sheets[ws.title] = [dict(zip(rows[0], r)) for r in rows[1:]] if rows else []
        wb.close()
        return sheets


def restaurant(igv: str, total: str) -> list[str]:
    return [
        "POLLERIA DEMO S.A.C.",
        f"RUC: {EM}",
        "FACTURA ELECTRONICA",
        "F001-00000321",
        "Fecha de Emisión: 15/03/2026",
        f"Señor(es): COMPRADORA S.A. RUC: {RE}",
        "Por concepto de: consumo de alimentos",
        "Op. Gravada: S/ 1000.00",
        f"IGV: S/ {igv}",
        f"Importe Total: S/ {total}",
    ]


class TestCorridaCompletaConConsultaRuc(unittest.TestCase):
    def _run(self, ciiu: str) -> dict[str, list[dict]]:
        with tempfile.TemporaryDirectory() as tmp:
            path = consulta_txt(Path(tmp), [f"{EM}|POLLERIA DEMO S.A.C.|ACTIVO|HABIDO|15/02/2015|{ciiu}||"])
            return scan({"rest": restaurant("105.00", "1105.00")}, "--consulta-ruc", str(path))

    def test_restaurante_con_ciiu_incluido_sale_de_revision(self) -> None:
        sheets = self._run("5610")
        row = sheets["cdp_detectados"][0]
        self.assertEqual(row["control_tasa_igv"], "tasa_especial_valida")
        self.assertIn("5610", row["ciiu_emisor"])
        self.assertIn("rucs_output.txt", row["ciiu_emisor"])
        kinds = [r["tipo_hallazgo"] for r in sheets["revision"] if r["id_documento"] == row["id_documento"]]
        self.assertNotIn("tasa_especial_verificar", kinds)
        self.assertNotIn("tasa_igv_atipica", kinds)

    def test_giro_no_incluido_va_a_revision_como_atipica(self) -> None:
        sheets = self._run("4711")
        row = sheets["cdp_detectados"][0]
        self.assertEqual(row["control_tasa_igv"], "tasa_igv_atipica")
        kinds = [r["tipo_hallazgo"] for r in sheets["revision"] if r["id_documento"] == row["id_documento"]]
        self.assertIn("tasa_igv_atipica", kinds)
        self.assertEqual(str(row["requiere_revision_humana"]), "True")

    def test_sin_consulta_queda_para_verificar(self) -> None:
        sheets = scan({"rest": restaurant("105.00", "1105.00")})
        row = sheets["cdp_detectados"][0]
        self.assertEqual(row["control_tasa_igv"], "tasa_especial_verificar")
        self.assertIn(row["ciiu_emisor"], (None, ""))


def padron_ruc_db(path: Path, rows: list[tuple[str, str, str, str, str]]) -> Path:
    """Base con el esquema exacto de consulta_ruc_masivo/padron_ruc.py: WAL, WITHOUT ROWID y sin padron_metadata."""
    conn = sqlite3.connect(path)
    try:
        conn.execute("PRAGMA journal_mode = WAL")
        conn.executescript("""
            CREATE TABLE padron (
                ruc            TEXT PRIMARY KEY,
                razon_social   TEXT,
                estado         TEXT,
                cond_domicilio TEXT,
                ubigeo         TEXT
            ) WITHOUT ROWID;
        """)
        conn.executemany("INSERT INTO padron VALUES (?,?,?,?,?)", rows)
        conn.commit()
        conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
    finally:
        conn.close()
    # padron_ruc.py no guarda la fecha de corte: se toma la del archivo. Se fija una antigua a propósito.
    old = time.mktime((2025, 7, 1, 12, 0, 0, 0, 0, -1))
    os.utime(path, (old, old))
    return path


class TestPadronSinLimiteDeAntiguedad(unittest.TestCase):
    def test_valor_por_defecto_es_sin_limite(self) -> None:
        self.assertIsNone(AppConfig().padron_max_age_days)
        self.assertIsNone(load_config(None).padron_max_age_days)


class TestPadronDeConsultaRucMasivo(unittest.TestCase):
    def test_lee_la_base_de_padron_ruc_py_aunque_sea_antigua(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            db = padron_ruc_db(Path(tmp) / "padron_ruc.db", [(EM, "POLLERIA DEMO S.A.C.", "ACTIVO", "HABIDO", "150101")])
            report = LocalPadronAdapter(db, max_age_days=None).evaluate_rucs([EM, OT])
        self.assertEqual(report.db_status, "disponible")
        self.assertEqual(report.cutoff_date, "2025-07-01")
        self.assertEqual(report.entries[EM].razon_social, "POLLERIA DEMO S.A.C.")
        self.assertEqual(report.entries[EM].status, "presente")
        self.assertEqual(report.entries[OT].status, "ausente")

    def test_lectura_no_deja_archivos_wal_ni_modifica_la_base(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            db = padron_ruc_db(Path(tmp) / "padron_ruc.db", [(EM, "DEMO S.A.C.", "ACTIVO", "HABIDO", "150101")])
            before = db.stat().st_mtime
            LocalPadronAdapter(db).evaluate_rucs([EM])
            self.assertEqual(db.stat().st_mtime, before)

    def test_resumen_muestra_la_fecha_de_corte_sin_bloquear(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            db = padron_ruc_db(Path(tmp) / "padron_ruc.db", [(EM, "POLLERIA DEMO S.A.C.", "ACTIVO", "HABIDO", "150101"),
                                                             (RE, "COMPRADORA S.A.", "ACTIVO", "HABIDO", "150101")])
            sheets = scan({"rest": restaurant("180.00", "1180.00")}, "--padron-db", str(db))
        resumen = {r["indicador"]: r["valor"] for r in sheets["resumen"]}
        self.assertIn("2025-07-01", str(resumen["Padrón RUC: fecha de corte"]))
        row = sheets["cdp_detectados"][0]
        kinds = [r["tipo_hallazgo"] for r in sheets["revision"] if r["id_documento"] == row["id_documento"]]
        self.assertFalse([k for k in kinds if "padron" in k or "obsoleto" in k], kinds)


def text_page(lines: list[str]) -> PageRead:
    layout = [LayoutLine(line_number=i, text=t, block_number=1, bbox=(10.0, i * 15.0, 500.0, i * 15.0 + 12),
                         source="native_pdf_text") for i, t in enumerate(lines, 1)]
    return PageRead(page_number=1, text="\n".join(lines), source="native_pdf_text", status="ok",
                    quality={"char_count": 100, "line_count": len(lines)}, layout_lines=layout, qr_payloads=[])


def column_page(rows: list[tuple[str, str | None]], *, values_first: bool = False) -> PageRead:
    """Rótulos en x=28 y valores en x=85 a la misma altura; con values_first el valor llega antes en la lectura."""
    layout, n = [], 0
    for i, (label, value) in enumerate(rows):
        y = 30.0 + 10 * i
        cells = [(28.0, 60.0, label), (85.0, 300.0, value)]
        if values_first:
            cells.reverse()
        for x0, x1, text in cells:
            if text is None:
                continue
            n += 1
            layout.append(LayoutLine(line_number=n, text=text, block_number=1, bbox=(x0, y, x1, y + 8), source="native_pdf_text"))
    return PageRead(page_number=1, text="\n".join(l.text for l in layout), source="native_pdf_text", status="ok",
                    quality={"char_count": 100, "line_count": len(layout)}, layout_lines=layout, qr_payloads=[])


def invoice_with_label(label: str) -> list[str]:
    return ["COMERCIAL ANDINA S.A.C.", f"RUC: {EM}", "FACTURA ELECTRONICA", "F001-00001234",
            "Fecha de Emisión: 15/03/2026", f"{label} COMPRADORA S.A.", f"RUC: {RE}",
            "Por concepto de: servicio de mantenimiento",
            "Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00"]


def recipient_comment(result) -> str:
    return " ".join(str(e.decision_comment) for e in result.field_evidence.get("recipient_ruc", []))


class TestRotuloAproximado(unittest.TestCase):
    def test_vocabulario_con_errores_de_ocr(self) -> None:
        for raw, canon in (("Senor(es):", "SEÑOR(ES)"), ("SEÑ0RES:", "SEÑOR(ES)"), ("Clíente :", "CLIENTE"),
                           ("Señore s:", "SEÑOR(ES)"), ("CLLIENTE:", "CLIENTE"), ("Adquiriente", "ADQUIRENTE"),
                           ("DESTINATARI0", "DESTINATARIO")):
            with self.subTest(raw):
                hit = fuzzy_recipient_label(raw)
                self.assertIsNotNone(hit)
                self.assertEqual(hit[0], canon)

    def test_palabras_parecidas_no_son_rotulo(self) -> None:
        for raw in ("SERIE:", "SEÑAL", "CLIENTES FRECUENTES GANAN PUNTOS", "Total", "RECEPCION", "SENATI", "USUARIOS 3"):
            with self.subTest(raw):
                self.assertIsNone(fuzzy_recipient_label(raw))

    def test_rotulo_solo_al_comienzo_de_la_linea(self) -> None:
        self.assertIsNone(fuzzy_recipient_label("Gracias por su preferencia estimado CLLIENTE"))

    def test_ruc_objetivo_bajo_rotulo_con_errores(self) -> None:
        for label in ("SEÑ0RES:", "Clíente:", "Señore s:", "CLLIENTE:"):
            with self.subTest(label):
                r = parse_document([text_page(invoice_with_label(label))], target_recipient_ruc=RE)[0]
                self.assertEqual(r.recipient_ruc, RE, r.observations)
                self.assertEqual(r.recipient_role_status, "verified_recipient")
                self.assertEqual(r.emitter_ruc, EM)
                self.assertIn("rótulo aproximado", recipient_comment(r))

    def test_rotulo_exacto_no_dice_aproximado(self) -> None:
        r = parse_document([text_page(invoice_with_label("Señor(es):"))], target_recipient_ruc=RE)[0]
        self.assertEqual(r.recipient_ruc, RE)
        self.assertNotIn("aproximado", recipient_comment(r))


class TestRotulosEnColumnas(unittest.TestCase):
    def rows(self, label: str) -> list[tuple[str, str | None]]:
        return [
            ("COMERCIAL ANDINA S.A.C.", None), (f"RUC: {EM}", None), ("FACTURA ELECTRONICA", None), ("F001-00001234", None),
            ("Fecha", ": 15/03/2026"), (label, ": COMPRADORA S.A."), ("R.U.C.", f": {RE}"), ("Dirección", ": Av. Lima 123"),
            ("Por concepto de: servicio de mantenimiento", None),
            ("Op. Gravada: S/ 100.00", None), ("IGV 18%: S/ 18.00", None), ("Importe Total: S/ 118.00", None),
        ]

    def test_columnas_con_valores_antes_o_despues_del_rotulo(self) -> None:
        for values_first in (False, True):
            for label in ("SEÑ0RES", "Clíente", "CLLIENTE", "Señor(es)"):
                with self.subTest(label=label, values_first=values_first):
                    r = parse_document([column_page(self.rows(label), values_first=values_first)], target_recipient_ruc=RE)[0]
                    self.assertEqual(r.recipient_ruc, RE, r.observations)
                    self.assertEqual(r.recipient_role_status, "verified_recipient")
                    self.assertEqual(r.emitter_ruc, EM)


class TestCasoNegativoVentaPropia(unittest.TestCase):
    def test_ruc_objetivo_en_el_membrete_no_se_toma_como_receptor(self) -> None:
        lines = ["COMPRADORA S.A.", f"RUC: {RE}", "FACTURA ELECTRONICA", "F001-00000077", "Fecha de Emisión: 15/03/2026",
                 "CLLIENTE: CLIENTE X S.A.", f"RUC: {OT}", "Por concepto de: venta de servicios",
                 "Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00"]
        r = parse_document([text_page(lines)], target_recipient_ruc=RE)[0]
        self.assertEqual(r.emitter_ruc, RE)
        self.assertEqual(r.recipient_ruc, OT)
        self.assertEqual(r.recipient_role_status, "mismatch")


if __name__ == "__main__":
    unittest.main()


REAL_HEADER = ("NumeroRuc|Nombre ó RazonSocial|Tipo de Contribuyente|Profesion u Oficio|Nombre Comercial|"
               "Condicion del Contribuyente|Estado del Contribuyente|Fecha de Inscripcion|Fecha de Inicio de Actividades|"
               "Departamento|Provincia|Distrito|Direccion|Telefono|Fax|Actividad de Comercio Exterior|Principal- CIIU|"
               "Secundario 1- CIIU|Secundario 2- CIIU|Afecto Nuevo RUS|Buen Contribuyente|Agente de Retencion|"
               "Agente de Percepcion VtaInt|Agente de Percepcion ComLiq")


def real_row(principal: str, sec1: str = "-", sec2: str = "-", estado: str = "ACTIVO", inicio: str = "01/10/2016") -> str:
    return (f"{EM}|DEMO S.A.C.|SOCIEDAD ANONIMA CERRADA|-|-|HABIDO|{estado}|23/09/2016|{inicio}|LIMA|LIMA|LIMA|AV. X 1|-|-|"
            f"SIN ACTIVIDAD|{principal}|{sec1}|{sec2}|NO|NO|NO|NO|NO")


class TestFormatoRealDeConsultaRuc(unittest.TestCase):
    """Formato verificado con una salida real de consulta_ruc.py: el CIIU viene como descripción Rev. 3, sin código."""

    def load(self, row: str):
        with tempfile.TemporaryDirectory() as tmp:
            return load_consulta_ruc(consulta_txt(Path(tmp), [row], header=REAL_HEADER)).activities[EM]

    def test_lee_descripciones_sin_codigo(self) -> None:
        act = self.load(real_row("RESTAURANTES, BARES Y CANTINAS.", "OTROS TIPOS DE VENTA AL POR MENOR."))
        self.assertEqual(act.principal, "RESTAURANTES, BARES Y CANTINAS")
        self.assertEqual(act.secondary, ("OTROS TIPOS DE VENTA AL POR MENOR",))
        self.assertEqual((act.estado, act.inicio_actividades), ("ACTIVO", "2016-10-01"))

    def test_giro_ajeno_en_descripcion_rev3_es_atipica(self) -> None:
        act = self.load(real_row("TRANSPORTE  DE CARGA POR CARRETERA."))
        chk = special("2026", "105.00", act=act)
        self.assertEqual(chk["status"], "tasa_igv_atipica")
        self.assertIn("el giro del emisor no es restaurante, hotel ni alojamiento", chk["reason"])

    def test_restaurantes_bares_y_cantinas_queda_para_verificar_con_motivo(self) -> None:
        for desc, word in (("RESTAURANTES, BARES Y CANTINAS.", "bares"), ("HOTELES, CAMPAMENTOS Y OTROS.", "campamentos")):
            with self.subTest(desc):
                chk = special("2026", "105.00", act=self.load(real_row(desc)))
                self.assertEqual(chk["status"], "tasa_especial_verificar")
                self.assertIn("Rev. 3", chk["reason"])
                self.assertIn(word, chk["reason"])

    def test_descripcion_rev4_del_regimen_decide(self) -> None:
        chk = special("2026", "105.00", act=self.load(real_row("ACTIVIDADES DE RESTAURANTES Y DE SERVICIO MÓVIL DE COMIDAS")))
        self.assertEqual(chk["status"], "tasa_especial_valida")
        chk = special("2026", "105.00", act=self.load(real_row("SUMINISTRO DE COMIDAS POR ENCARGO")))
        self.assertEqual(chk["status"], "tasa_igv_atipica")

    def test_restaurante_solo_secundario_queda_para_verificar(self) -> None:
        chk = special("2026", "105.00", act=self.load(real_row("VENTA AL POR MENOR COMBUSTIBLES.", "RESTAURANTES, BARES Y CANTINAS.")))
        self.assertEqual(chk["status"], "tasa_especial_verificar")
        self.assertIn("secundaria", chk["reason"])

    def test_codigo_rev3_de_hotel_o_restaurante_queda_para_verificar_y_otro_es_atipica(self) -> None:
        self.assertEqual(special("2026", "105.00", act=activity("55205"))["status"], "tasa_especial_verificar")
        self.assertEqual(special("2026", "105.00", act=activity("60230"))["status"], "tasa_igv_atipica")
        self.assertEqual(special("2026", "105.00", act=activity("5520"))["status"], "tasa_especial_verificar")


def make_input(tmp: Path, docs: dict[str, list[str]]) -> Path:
    inp = tmp / "entrada"
    inp.mkdir()
    for name, lines in docs.items():
        doc = pymupdf.open()
        pdf_page = doc.new_page(width=595, height=842)
        for i, line in enumerate(lines):
            pdf_page.insert_text((36, 40 + 14 * i), line, fontsize=10, fontname="helv")
        doc.save(inp / f"{name}.pdf")
        doc.close()
    return inp


def sheets_of(out: Path) -> dict[str, list[dict]]:
    wb = openpyxl.load_workbook(out / "CDP.xlsx", read_only=True)
    try:
        result = {}
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            result[ws.title] = [dict(zip(rows[0], r)) for r in rows[1:]] if rows else []
        return result
    finally:
        wb.close()


def cli(*args: str) -> int:
    with contextlib.redirect_stdout(io.StringIO()):
        return cli_main(list(args))


class TestLecturaCompletaGuardada(unittest.TestCase):
    def test_por_defecto_se_guarda_con_lineas_coordenadas_y_confianza(self) -> None:
        self.assertTrue(AppConfig().keep_full_reading)
        self.assertTrue(load_config(None).keep_full_reading)
        from ocr_fiscal.reading_cache import ReadingCache

        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            inp = make_input(tmp, {"rest": restaurant("180.00", "1180.00")})
            self.assertEqual(cli("scan", "--input", str(inp), "--output", str(tmp / "out"), "--recipient-ruc", RE), 0)
            cache = ReadingCache(tmp / "out" / ".ocr_fiscal" / "lecturas.sqlite3")
            self.assertEqual(cache.count(), 1)
            read = cache.first()
            lines = [line for page in read.pages for line in page.layout_lines]
            self.assertTrue(any(f"RUC: {EM}" in line.text for line in lines))
            self.assertTrue(all(line.bbox is not None for line in lines))
            cache.close()

    def test_opcion_para_no_guardar(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            inp = make_input(tmp, {"rest": restaurant("180.00", "1180.00")})
            (tmp / "cfg.json").write_text('{"guardar_lectura_completa": false}', encoding="utf-8")
            cli("scan", "--input", str(inp), "--output", str(tmp / "out"), "--recipient-ruc", RE, "--config", str(tmp / "cfg.json"))
            self.assertFalse((tmp / "out" / ".ocr_fiscal" / "lecturas.sqlite3").exists())
            self.assertTrue((tmp / "out" / "CDP.xlsx").is_file())


class TestRedetectar(unittest.TestCase):
    def test_redetectar_reproduce_el_excel_sin_leer_los_documentos(self) -> None:
        from unittest import mock

        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            inp = make_input(tmp, {"rest": restaurant("105.00", "1105.00"), "otra": invoice_with_label("Señor(es):")})
            out = tmp / "out"
            self.assertEqual(cli("scan", "--input", str(inp), "--output", str(out), "--recipient-ruc", RE), 0)
            with mock.patch("ocr_fiscal.processor.read_work_unit", side_effect=AssertionError("no debe leer")):
                self.assertEqual(cli("redetectar", "--output", str(out), "--destino", str(tmp / "re")), 0)
            before, after = sheets_of(out), sheets_of(tmp / "re")
            keys = ("ruc_emisor", "serie", "correlativo", "ruc_adquirente", "igv", "total_leido", "control_tasa_igv")
            self.assertEqual([{k: r[k] for k in keys} for r in before["cdp_detectados"]],
                             [{k: r[k] for k in keys} for r in after["cdp_detectados"]])
            self.assertEqual(sorted(r["tipo_hallazgo"] for r in before["revision"]),
                             sorted(r["tipo_hallazgo"] for r in after["revision"]))

    def test_redetectar_aplica_reglas_nuevas(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            inp = make_input(tmp, {"rest": restaurant("105.00", "1105.00")})
            out = tmp / "out"
            cli("scan", "--input", str(inp), "--output", str(out), "--recipient-ruc", RE)
            consulta = consulta_txt(tmp, [f"{EM}|POLLERIA DEMO S.A.C.|ACTIVO|HABIDO|15/02/2015|5610||"])
            self.assertEqual(cli("redetectar", "--output", str(out), "--destino", str(tmp / "re"), "--consulta-ruc", str(consulta)), 0)
            self.assertEqual(sheets_of(out)["cdp_detectados"][0]["control_tasa_igv"], "tasa_especial_verificar")
            self.assertEqual(sheets_of(tmp / "re")["cdp_detectados"][0]["control_tasa_igv"], "tasa_especial_valida")

    def test_documento_sin_lectura_guardada_va_a_revision(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            inp = make_input(tmp, {"rest": restaurant("180.00", "1180.00")})
            out = tmp / "out"
            cli("scan", "--input", str(inp), "--output", str(out), "--recipient-ruc", RE)
            # Un documento que llega después de la corrida de origen no tiene lectura guardada.
            (tmp / "x").mkdir()
            extra = make_input(tmp / "x", {"nuevo": invoice_with_label("Cliente:")})
            (inp / "nuevo.pdf").write_bytes((extra / "nuevo.pdf").read_bytes())
            self.assertEqual(cli("redetectar", "--output", str(out), "--destino", str(tmp / "re")), 0)
            sheets = sheets_of(tmp / "re")
            self.assertIn("lectura no guardada", json_dump(sheets["revision"]) + json_dump(sheets["inventario"]))
            self.assertEqual(len(sheets["cdp_detectados"]), 1)

    def test_sin_lecturas_guardadas_se_niega(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            inp = make_input(tmp, {"rest": restaurant("180.00", "1180.00")})
            (tmp / "cfg.json").write_text('{"guardar_lectura_completa": false}', encoding="utf-8")
            cli("scan", "--input", str(inp), "--output", str(tmp / "out"), "--recipient-ruc", RE, "--config", str(tmp / "cfg.json"))
            with contextlib.redirect_stderr(io.StringIO()) as err:
                self.assertEqual(cli("redetectar", "--output", str(tmp / "out"), "--destino", str(tmp / "re")), 2)
            self.assertIn("lecturas", err.getvalue())


def json_dump(value) -> str:
    import json

    return json.dumps(value, ensure_ascii=False, default=str)


def parse_lines(lines: list[str]):
    return parse_document([text_page(lines)], target_recipient_ruc=RE)


class TestSerieNumericaYFacturaEnIngles(unittest.TestCase):
    def test_serie_numerica_con_n_despues_del_guion(self) -> None:
        r = parse_lines(["COMERCIAL ANDINA S.A.C.", f"R.U.C. {EM}", "FACTURA", "001 - N° 0012345", "Lima, 15 de marzo de 2026",
                         f"Señor(es): COMPRADORA S.A. R.U.C. {RE}", "SUB TOTAL 100.00", "I.G.V. 18.00", "TOTAL S/ 118.00"])[0]
        self.assertEqual(r.document_type, "01")
        self.assertIn(r.series, ("001", "0001"))
        self.assertEqual(r.correlativo_normalized, "00012345")

    def test_telefono_con_guion_sigue_sin_ser_serie(self) -> None:
        r = parse_lines(["COMERCIAL ANDINA S.A.C.", f"R.U.C. {EM}", "Telf. 044 - 223344", "FACTURA ELECTRONICA", "F001-00000077",
                         "Fecha de Emisión: 15/03/2026", f"Señor(es): COMPRADORA S.A. R.U.C. {RE}",
                         "Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00"])[0]
        self.assertEqual((r.series, r.correlativo_normalized), ("F001", "00000077"))

    def test_factura_electronica_en_ingles(self) -> None:
        r = parse_lines(["ANDEAN TRADING S.A.C.", f"RUC: {EM}", "ELECTRONIC INVOICE", "F001-00001234", "Issue date: 15/03/2026",
                         f"Customer: COMPRADORA S.A. RUC: {RE}", "Taxable amount: 100.00", "VAT 18%: 18.00", "Total amount: 118.00"])[0]
        self.assertEqual(r.detection_status, DETECTED)
        self.assertEqual((r.document_type, r.series, r.correlativo_normalized), ("01", "F001", "00001234"))
        self.assertEqual((r.emitter_ruc, r.recipient_ruc, r.issue_date), (EM, RE, "2026-03-15"))
        self.assertEqual((r.gravado, r.igv, r.total), ("100.00", "18.00", "118.00"))

    def test_commercial_invoice_de_importacion_no_es_factura(self) -> None:
        r = parse_lines(["GLOBAL SUPPLIES INC.", "COMMERCIAL INVOICE", "No. 2026-0045", "Date: 03/15/2026",
                         "Bill to: COMPRADORA S.A.", "Total amount: USD 1,180.00"])[0]
        self.assertNotEqual(r.document_type, "01")
        self.assertNotEqual(r.detection_status, DETECTED)

    def test_ticket_sin_tipo_confirmado_sigue_indeterminado(self) -> None:
        r = parse_lines(["MINIMARKET ANDINO", f"RUC {EM}", "TICKET NRO 0001-000123", "15/03/2026 12:30", "GASEOSA 1 5.00",
                         "TOTAL S/ 5.00", "EFECTIVO 10.00", "VUELTO 5.00"])[0]
        self.assertNotEqual(r.detection_status, DETECTED)
        self.assertTrue(r.requires_human_review)


class TestCiiuMixtoSeDecidePorLasDescripcionesDelComprobante(unittest.TestCase):
    """Con «RESTAURANTES, BARES Y CANTINAS» u «HOTELES, CAMPAMENTOS Y OTROS»
    deciden las descripciones del propio comprobante."""

    def act(self, desc: str) -> EmitterActivity:
        return activity(desc)

    def decide(self, ciiu: str, desc: str | None) -> dict:
        chk = evaluate_igv_rate(R(document_type="01", gravado="1000.00", igv="105.00", issue_date="2026-03-01"))
        return refine_special_rate_check(chk, self.act(ciiu), issue_date="2026-03-01", document_type="01", descriptions=desc)

    def test_comida_en_restaurante_es_valida(self) -> None:
        for desc in ("CHULETA DE ALPACA A LA PLANCHA | AGUA C/S GAS", "MENU EJECUTIVO", "consumo de alimentos",
                     "POLLO A LA BRASA 1/4 | GASEOSA 500ML", "CEVICHE MIXTO"):
            with self.subTest(desc):
                chk = self.decide("RESTAURANTES, BARES Y CANTINAS", desc)
                self.assertEqual(chk["status"], "tasa_especial_valida")
                self.assertIn("descripciones del comprobante", chk["reason"])

    def test_solo_bebidas_de_bar_queda_para_verificar(self) -> None:
        chk = self.decide("RESTAURANTES, BARES Y CANTINAS", "CERVEZA CUSQUEÑA 620ML | PISCO SOUR | COPA DE VINO")
        self.assertEqual(chk["status"], "tasa_especial_verificar")
        self.assertIn("bar", chk["reason"])

    def test_alojamiento_en_hotel_es_valido_y_campamento_no(self) -> None:
        for desc in ("HOSPEDAJE 2 NOCHES HABITACION DOBLE", "Alojamiento del 12/01 al 14/01", "HABITACION MATRIMONIAL"):
            with self.subTest(desc):
                self.assertEqual(self.decide("HOTELES, CAMPAMENTOS Y OTROS", desc)["status"], "tasa_especial_valida")
        self.assertEqual(self.decide("HOTELES, CAMPAMENTOS Y OTROS", "ZONA DE CAMPING CARPA 2 NOCHES")["status"],
                         "tasa_especial_verificar")

    def test_sin_descripcion_o_ajena_queda_para_verificar(self) -> None:
        for desc in (None, "", "SERVICIO DE MANTENIMIENTO", "ALQUILER DE SALON"):
            with self.subTest(desc):
                chk = self.decide("RESTAURANTES, BARES Y CANTINAS", desc)
                self.assertEqual(chk["status"], "tasa_especial_verificar")

    def test_comida_en_hotel_no_basta_y_alojamiento_en_restaurante_no_basta(self) -> None:
        self.assertEqual(self.decide("HOTELES, CAMPAMENTOS Y OTROS", "CEVICHE MIXTO")["status"], "tasa_especial_verificar")
        self.assertEqual(self.decide("RESTAURANTES, BARES Y CANTINAS", "HOSPEDAJE 1 NOCHE")["status"], "tasa_especial_verificar")

    def test_corrida_completa_lee_la_descripcion_del_comprobante(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = consulta_txt(Path(tmp), [real_row("RESTAURANTES, BARES Y CANTINAS.")], header=REAL_HEADER)
            lines = restaurant("105.00", "1105.00")
            sheets = scan({"rest": lines}, "--consulta-ruc", str(path))
        self.assertEqual(sheets["cdp_detectados"][0]["control_tasa_igv"], "tasa_especial_valida")


def cells_page(cells: list[tuple[str, float, float]], *, source: str = "native_pdf_text") -> PageRead:
    """Página con líneas sueltas en (x, y), en el orden dado. Cada línea mide 10 de alto y 5.5 por carácter."""
    layout = [LayoutLine(line_number=i, text=t, block_number=1, bbox=(x, y, x + 5.5 * len(t), y + 10), source=source)
              for i, (t, x, y) in enumerate(cells, 1)]
    return PageRead(page_number=1, text="\n".join(t for t, _, _ in cells), source=source, status="ok",
                    quality={"char_count": 100, "line_count": len(cells)}, layout_lines=layout, qr_payloads=[])


INVOICE_TOP = [("COMERCIAL ANDINA S.A.C.", 40, 20), (f"RUC: {EM}", 40, 32), ("FACTURA ELECTRONICA", 300, 20),
               ("F001-00000421", 300, 32), ("Fecha de Emisión: 15/03/2026", 40, 44),
               (f"Señor(es): COMPRADORA S.A. RUC: {RE}", 40, 56)]


def items_of(r) -> list[tuple]:
    return [(li.code, li.description_normalized, li.quantity_normalized, li.line_total_normalized) for li in r.line_items]


class TestDescripcionesEItems(unittest.TestCase):
    def test_importe_en_celda_aparte_de_la_misma_fila(self) -> None:
        r = parse_document([cells_page(INVOICE_TOP + [
            ("CANT. UM CÓD. DESCRIPCIÓN V/U P/U IMPORTE", 50, 100),
            ("3 NIU 100-003 MOTOBOMBA GASOLINA 3.5 HP 313.560 370.001", 59, 115), ("1,110.00", 527, 115),
            ("1 ZZ T&L0020 OUTSOURCING TRIBUTARIO", 59, 127), ("200.000 236.000", 409, 127), ("236.00", 534, 127),
            ("Op. Gravada: S/ 1,140.68", 300, 200), ("IGV 18%: S/ 205.32", 300, 212), ("Importe Total: S/ 1,346.00", 300, 224),
        ])], target_recipient_ruc=RE)[0]
        self.assertEqual(items_of(r), [("100-003", "MOTOBOMBA GASOLINA 3.5 HP", "3.00", "1110.00"),
                                       ("T&L0020", "OUTSOURCING TRIBUTARIO", "1.00", "236.00")])
        self.assertEqual(r.general_description, "MOTOBOMBA GASOLINA 3.5 HP | OUTSOURCING TRIBUTARIO")

    def test_segunda_linea_de_cabecera_no_es_descripcion(self) -> None:
        r = parse_document([cells_page(INVOICE_TOP + [
            ("PRECIO VALOR DE", 438, 95), ("DESCRIPCIÓN", 184, 100), ("UNITARIO VENTA", 436, 104),
            ("POR EL APORTE ESTABLECIDO EN LA LEY 28749", 40, 115),
            ("Importe Total: S/ 543.94", 300, 224),
        ])], target_recipient_ruc=RE)[0]
        self.assertNotIn("UNITARIO", r.general_description or "")
        self.assertIn("POR EL APORTE ESTABLECIDO", r.general_description or "")

    def test_pie_y_lineas_lejanas_no_continuan_el_item(self) -> None:
        r = parse_document([cells_page(INVOICE_TOP + [
            ("CODIGO DESCRIPCION UNIDAD CANTIDAD VALOR UNITARIO VALOR DE VENTA", 40, 100),
            ("0006 LER LEY 28749 Clientes Libres NIU 1.00 300,822.06 300,822.06", 40, 115),
            ("Fecha de Vencimiento : 29/01/2019", 60, 400),
            ("TRESCIENTOS MIL OCHOCIENTOS VEINTIDOS CON 06/100 SOLES", 60, 412),
            ("Importe Total: S/ 300,822.06", 300, 500),
        ])], target_recipient_ruc=RE)[0]
        self.assertEqual(r.general_description, "LER LEY 28749 Clientes Libres")
        self.assertEqual(items_of(r)[0][1], "LER LEY 28749 Clientes Libres")

    def test_segunda_linea_cercana_si_continua_el_item(self) -> None:
        r = parse_document([cells_page(INVOICE_TOP + [
            ("CODIGO DESCRIPCION UNIDAD CANTIDAD VALOR UNITARIO VALOR DE VENTA", 40, 100),
            ("0001 Venta Energia CL NIU 1.00 1,463,452.91 1,463,452.91", 48, 115),
            ("0001 Venta de Potencia CL / FACTURACION DE ENERGIA Y POTENCIA - NIU 1.00 60,709.00", 48, 125), ("60,709.00", 545, 125),
            ("MAYO 2019.", 88, 135),
            ("Importe Total: S/ 1,524,161.91", 300, 300),
        ])], target_recipient_ruc=RE)[0]
        self.assertEqual([i[1] for i in items_of(r)],
                         ["Venta Energia CL", "Venta de Potencia CL / FACTURACION DE ENERGIA Y POTENCIA - MAYO 2019"])
        self.assertEqual(items_of(r)[1][3], "60709.00")


def numbered(page: PageRead, number: int) -> PageRead:
    page.page_number = number
    return page


class TestCopiaDelComprobanteEnOtraPagina(unittest.TestCase):
    BODY = [("CODIGO DESCRIPCION UNIDAD CANTIDAD VALOR UNITARIO VALOR DE VENTA", 40, 100),
            ("0001 SERVICIO DE ALMACENAJE NIU 1.00 1,000.00 1,000.00", 40, 115),
            ("Op. Gravada: S/ 1,000.00", 300, 200), ("IGV 18%: S/ 180.00", 300, 212), ("Importe Total: S/ 1,180.00", 300, 224)]

    def test_copia_repetida_no_duplica_los_items(self) -> None:
        pages = [numbered(cells_page(INVOICE_TOP + self.BODY), 1), numbered(cells_page(INVOICE_TOP + self.BODY), 2)]
        results = parse_document(pages, target_recipient_ruc=RE)
        self.assertEqual(len(results), 1)
        self.assertEqual([i[1] for i in items_of(results[0])], ["SERVICIO DE ALMACENAJE"])

    def test_pagina_de_continuacion_con_items_distintos_los_conserva(self) -> None:
        second = [("CODIGO DESCRIPCION UNIDAD CANTIDAD VALOR UNITARIO VALOR DE VENTA", 40, 100),
                  ("0002 SERVICIO DE DISTRIBUCION NIU 1.00 500.00 500.00", 40, 115),
                  ("0001 SERVICIO DE ALMACENAJE NIU 1.00 1,000.00 1,000.00", 40, 127)]
        pages = [numbered(cells_page(INVOICE_TOP + self.BODY[:2]), 1), numbered(cells_page(INVOICE_TOP + second + self.BODY[2:]), 2)]
        result = parse_document(pages, target_recipient_ruc=RE)[0]
        self.assertEqual(sorted(i[1] for i in items_of(result)),
                         ["SERVICIO DE ALMACENAJE", "SERVICIO DE ALMACENAJE", "SERVICIO DE DISTRIBUCION"])


class TestCodigoDelItem(unittest.TestCase):
    def test_palabra_sin_digitos_no_es_codigo(self) -> None:
        r = parse_document([cells_page(INVOICE_TOP + [
            ("CANT. UM CÓD. DESCRIPCIÓN V/U P/U IMPORTE", 50, 100),
            ("1.0 UN POR SERVICIOS DE SISTEMA DE DECISION 180.000 180.00", 59, 115),
            ("Op. Gravada: S/ 180.00", 300, 200), ("IGV 18%: S/ 32.40", 300, 212), ("Importe Total: S/ 212.40", 300, 224),
        ])], target_recipient_ruc=RE)[0]
        self.assertEqual(items_of(r), [(None, "POR SERVICIOS DE SISTEMA DE DECISION", "1.00", "180.00")])


class TestFilaConCantidadYPrecioPorUnidad(unittest.TestCase):
    def test_cantidad_unidad_y_precio_por_unidad_no_entran_en_la_descripcion(self) -> None:
        r = parse_document([cells_page(INVOICE_TOP + [
            ("PRECIO VALOR DE", 438, 95), ("DESCRIPCIÓN", 184, 100), ("UNITARIO VENTA", 436, 104),
            ("Aporte de los usuarios de electricidad para la Electrifcacion Rural 63 216,156 MWh 8,600 S//MWh", 40, 115),
            ("543 658,94", 600, 115),
            ("TOTAL PRESENTE FACTURA :", 40, 135), ("S/", 484, 135), ("543 658,94", 555, 135),
        ])], target_recipient_ruc=RE)[0]
        self.assertEqual(r.general_description, "Aporte de los usuarios de electricidad para la Electrifcacion Rural")
        li = r.line_items[0]
        self.assertEqual((li.description_normalized, li.quantity_normalized, li.unit_normalized, li.line_total_normalized),
                         ("Aporte de los usuarios de electricidad para la Electrifcacion Rural", "63216.156", "MWH", "543658.94"))
        # Coma decimal en toda la fila: «8,600» es 8.600 por MWh, no 8 600 (63 216,156 × 8,600 = 543 658,94).
        self.assertEqual(li.unit_price_normalized, "8.600")
        self.assertNotIn("line_items", r.secondary_review_fields)


class TestFilaQueEmpiezaConCantidadYUnidad(unittest.TestCase):
    def test_unidad_dentro_de_la_descripcion_no_cambia_la_forma_de_la_fila(self) -> None:
        r = parse_document([cells_page(INVOICE_TOP + [
            ("CANT. UM CÓD. DESCRIPCIÓN V/U P/U IMPORTE", 50, 100),
            ("4 NIU 201-025D GENERADOR GASOLINA (X) TG3100CXR 31 KW 699.150 824.997", 59, 115), ("3,299.99", 600, 115),
            ("Op. Gravada: S/ 2,796.60", 300, 200), ("IGV 18%: S/ 503.39", 300, 212), ("Importe Total: S/ 3,299.99", 300, 224),
        ])], target_recipient_ruc=RE)[0]
        self.assertEqual(items_of(r), [("201-025D", "GENERADOR GASOLINA (X) TG3100CXR 31 KW", "4.00", "3299.99")])


class TestSerieDelRecuadroFiscal(unittest.TestCase):
    """La serie se perdía por candidatos falsos fuera del recuadro fiscal."""

    def build(self, box_serial: str, noise: list[str], *, second_box: str | None = None) -> list[str]:
        box = [box_serial, second_box] if second_box else [box_serial]
        lines = ["RUC " + EM, "FACTURA", "ELECTRÓNICA", "COMERCIAL ANDINA S.A.C.", *box,
                 "Teléfono: (511) 500-6777 / 627 -3737", f"Señor(es): COMPRADORA S.A. RUC: {RE}",
                 "Fecha de Emisión: 13/01/2026", *noise]
        return lines + ["Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00"]

    def test_la_serie_sola_junto_al_titulo_gana_a_telefono_cuenta_e_item(self) -> None:
        for serial, expected in (("F001-008511", ("F001", "00008511")), ("FFF1-000421", ("FFF1", "00000421")),
                                 ("F002-00000757", ("F002", "00000757"))):
            with self.subTest(serial):
                r = parse_lines(self.build(serial, ["CCI: 009-097-000-4594253", "5 NIU DESBROZADORA TOYAMA RT52LB - 516 CC 242.00 285.56 1,427.80"]))[0]
                self.assertEqual((r.series, r.correlativo_normalized), expected)
                self.assertTrue(any("recuadro fiscal" in o for o in r.observations), r.observations)

    def test_dos_series_en_recuadro_siguen_en_conflicto(self) -> None:
        r = parse_lines(self.build("F001-00000001", [], second_box="F002-00000002"))[0]
        self.assertIsNone(r.correlativo_normalized)
        self.assertIn("F002-00000002", r.candidates.get("document_identity", []))
        self.assertTrue(r.requires_human_review)

    def test_serie_numerica_de_ceros_no_es_serie(self) -> None:
        r = parse_lines(["COMERCIAL ANDINA S.A.C.", f"R.U.C. {EM}", "FACTURA", "001 - N° 0012345", "CCI 000-4594253",
                         f"Señor(es): COMPRADORA S.A. R.U.C. {RE}", "Fecha: 15/03/2026", "TOTAL S/ 118.00"])[0]
        self.assertEqual((r.series, r.correlativo_normalized), ("001", "00012345"))


class TestDescripcionYBaseValidadas(unittest.TestCase):
    def test_glosa_de_la_tabla_se_conserva_junto_al_item(self) -> None:
        r = parse_document([cells_page(INVOICE_TOP + [
            ("PRECIO VALOR DE", 438, 95), ("DESCRIPCIÓN", 184, 100), ("UNITARIO VENTA", 436, 104),
            ("POR EL APORTE ESTABLECIDO EN LA LEY 28749,", 40, 115), ("CORRESPONDIENTE AL MES DE JULIO 2020.", 40, 123),
            ("Aporte de los usuarios para la", 40, 131), ("34 926,71", 554, 131), ("Electrifcacion Rural", 40, 139),
            ("TOTAL PRESENTE FACTURA :", 40, 160), ("S/", 484, 160), ("34 926,71", 555, 160),
        ])], target_recipient_ruc=RE)[0]
        self.assertEqual(r.general_description,
                         "POR EL APORTE ESTABLECIDO EN LA LEY 28749, CORRESPONDIENTE AL MES DE JULIO 2020 | "
                         "Aporte de los usuarios para la Electrifcacion Rural")
        self.assertEqual([li.description_normalized for li in r.line_items], ["Aporte de los usuarios para la Electrifcacion Rural"])

    def test_servicio_que_abre_la_descripcion_no_es_unidad(self) -> None:
        r = parse_document([cells_page(INVOICE_TOP + [
            ("CANTIDAD DESCRIPCIÓN V. UNITARIO IMPORTE", 23, 100),
            ("1.00 SERVICIO DE ALMACENAJE - NOVIEMBRE 2025 14,937.50", 36, 115), ("14,937.50", 600, 115),
            ("Total Ope. Gravadas S/", 374, 200), ("14,937.50", 536, 200),
            ("IGV 18%: S/ 2,688.75", 300, 212), ("Importe Total: S/ 17,626.25", 300, 224),
        ])], target_recipient_ruc=RE)[0]
        self.assertEqual(r.general_description, "SERVICIO DE ALMACENAJE - NOVIEMBRE 2025")
        self.assertEqual(items_of(r), [(None, "SERVICIO DE ALMACENAJE - NOVIEMBRE 2025", "1.00", "14937.50")])
        self.assertEqual(r.gravado, "14937.50")

    def test_rotulo_ope_gravadas_en_linea(self) -> None:
        r = parse_lines(invoice_with_label("Señor(es):")[:-3] + ["Total Ope. Gravadas S/ 100.00", "IGV 18%: S/ 18.00",
                                                                 "Importe Total: S/ 118.00"])[0]
        self.assertEqual(r.gravado, "100.00")


class TestRotuloDeOperacionesAlFinalDeLinea(unittest.TestCase):
    def test_rotulo_al_final_de_una_linea_de_cuentas(self) -> None:
        r = parse_document([cells_page(INVOICE_TOP + [
            ("CANTIDAD DESCRIPCIÓN V. UNITARIO IMPORTE", 23, 100),
            ("1.00 SERVICIO DE ALMACENAJE - NOVIEMBRE 2025 14,937.50", 36, 115), ("14,937.50", 600, 115),
            ("BANCO X SOL CTA CTE: 0004567753 Total Ope. Gravadas S/", 26, 200), ("14,937.50", 536, 200),
            ("BANCO X USD CTA CTE: 0005296286 Total Ope. Inafectas S/", 26, 212), ("0.00", 554, 212),
            ("IGV 18%: S/ 2,688.75", 300, 224), ("Importe Total: S/ 17,626.25", 300, 236),
        ])], target_recipient_ruc=RE)[0]
        self.assertEqual((r.gravado, r.inafecto), ("14937.50", "0.00"))

    def test_total_suelto_al_final_de_una_linea_no_se_empareja(self) -> None:
        r = parse_document([cells_page(INVOICE_TOP + [
            ("Por concepto de: servicio de mantenimiento", 40, 100),
            ("Leyenda: el pago se hace al TOTAL", 26, 200), ("999.99", 536, 200),
            ("Op. Gravada: S/ 100.00", 300, 224), ("IGV 18%: S/ 18.00", 300, 236), ("Importe Total: S/ 118.00", 300, 248),
        ])], target_recipient_ruc=RE)[0]
        self.assertEqual(r.total, "118.00")


def ocr_page(lines: list[str]) -> PageRead:
    layout = [LayoutLine(line_number=i, text=t, block_number=1, bbox=(10.0, i * 15.0, 500.0, i * 15.0 + 12),
                         source="ocr", ocr_confidence=0.9) for i, t in enumerate(lines, 1)]
    return PageRead(page_number=1, text="\n".join(lines), source="ocr", status="ok",
                    quality={"char_count": 100, "line_count": len(lines)}, layout_lines=layout, qr_payloads=[])


class TestImportesConErroresDeOcr(unittest.TestCase):
    def ocr_invoice(self, totals: list[str]) -> list[str]:
        return ["ENERGIA DEMO S.A.A.", f"RUC: {EM}", "FACTURA ELECTRONICA", "F010-00017663", "Lima 14 de Enero del 2019",
                f"Señor(es) : COMPRADORA S.A.", f"R.U.C. : {RE}", "PEAJE CONEXION SPT 341,130.00", *totals]

    def test_igv_leido_16v_y_moneda_leida_si(self) -> None:
        r = parse_document([ocr_page(self.ocr_invoice([
            "TOTAL AFECTO S/ 642,389.16", "y TOTAL INAFECTO Si 0.00", "16V S 115,630.05",
            "Para consultar su comprobante ingrese a: OTROS CARGOS S/ 0.00",
            "Autorizado Mediante Resolución de Intendencia IMPORTE TOTAL — |S/ 758,019.21"]))],
            target_recipient_ruc=RE)[0]
        self.assertEqual((r.gravado, r.igv, r.total), ("642389.16", "115630.05", "758019.21"))
        self.assertEqual(r.inafecto, "0.00")

    def test_inafecto_con_si_e_igv_cero(self) -> None:
        r = parse_document([ocr_page(self.ocr_invoice([
            "TOTAL AFECTO S 0.00", "4 TOTAL INAFECTO S 300,822.06", "! 16V Si 0.00",
            "Autorizado Mediante Resolución de Intendencia IMPORTE TOTAL Si 300,822.06"]))],
            target_recipient_ruc=RE)[0]
        self.assertEqual((r.inafecto, r.igv, r.total), ("300822.06", "0.00", "300822.06"))

    def test_16v_dentro_de_una_descripcion_no_es_igv(self) -> None:
        from ocr_fiscal.amounts import amount_mentions

        self.assertFalse([m for m in amount_mentions("PILA ALCALINA 16V S/ 50.00") if m[0] == "igv"])
        self.assertFalse([m for m in amount_mentions("BATERIA 16V 120.50") if m[0] == "igv"])


class TestAjustesDeLaValidacion(unittest.TestCase):
    def test_serie_a_cuatro_lineas_del_titulo(self) -> None:
        lines = ["RUC " + EM, "FACTURA", "ELECTRÓNICA", "COMERCIAL ANDINA S.A.C.", "AV. REDUCTO 1310", "FFF1-000421",
                 f"RUC : {RE}", "FECHA EMISIÓN : 27/02/2026",
                 "5 NIU RT52L-B DESBROZADORA TOYAMA RT52LB - 516 CC 242.00 285.56 1,427.80",
                 "Op. Gravada: S/ 100.00", "IGV 18%: S/ 18.00", "Importe Total: S/ 118.00"]
        r = parse_lines(lines)[0]
        self.assertEqual((r.series, r.correlativo_normalized), ("FFF1", "00000421"))

    def test_continuacion_del_item_sin_cabecera_de_tabla(self) -> None:
        r = parse_document([ocr_page(["EMPRESA DEMO S.A.", f"R.U.C.: {EM}", "FACTURA ELECTRONICA", "Nº F001- 01121729",
                                      f"SEÑOR(ES) :COMPRADORA S.A. FECHA :2026-02-28", f"R.U.C. :{RE} MONEDA :SOLES",
                                      "1.0 UN POR SERVICIOS DE SISTEMA DE DECISION DURANTE EL PERIODO DE 01/02/2026 AL 180.000 180.00",
                                      "28/02/2026, MONTO MINIMO",
                                      "SON :DOSCIENTOS DOCE CON 40/100 SOLES."])], target_recipient_ruc=RE)[0]
        self.assertEqual(r.line_items[0].description_normalized,
                         "POR SERVICIOS DE SISTEMA DE DECISION DURANTE EL PERIODO DE 01/02/2026 AL 28/02/2026, MONTO MINIMO")


class TestGlosaSinBasuraDeOcr(unittest.TestCase):
    def test_linea_ilegible_no_entra_como_glosa(self) -> None:
        r = parse_document([ocr_page(["RESTAURANTE DEMO S.A.C.", f"R.U.C.: {EM}", "FACTURA ELECTRONICA", "F401-00009631",
                                      "Fecha Emision : 13/01/2026", f"Empresa: COMPRADORA S.A. RUC : {RE}",
                                      "Producto Cant. P.Unit Total", "soegeeewnwessoresssoosesesococessoceccce",
                                      "Ice Tea Tumbo 1.00 15.00 15.00", "Ensalada nikke 1.00 36.00 36.00",
                                      "Importe Total S/: 51.00"])], target_recipient_ruc=RE)[0]
        self.assertNotIn("soegee", r.general_description or "")
        self.assertIn("Ice Tea Tumbo", r.general_description or "")


class TestLoNoDetectado(unittest.TestCase):
    """Datos que antes no se detectaban."""

    def test_fila_cantidad_guion_descripcion_unidad_precio_y_valor_de_venta(self) -> None:
        r = parse_document([ocr_page(["OFICINAS DEMO S.A.", f"R.U.C. {EM}", "FACTURA ELECTRONICA", "F002-00000757",
                                      f"Nombre/Razón Social: COMPRADORA S.A. RUC: {RE}", "Fecha Emisión: 08/01/2026",
                                      "Cantidad Código Descripción Und. P. Unitario Valor Venta",
                                      "1 - Servicio Oficina Privada N°716 - Febrero NIU 1,274.40 1,080.00",
                                      "Op. Gravada US$ 1,080.00", "I.G.V US$ 194.40", "Importe Total US$ 1,274.40"])],
                           target_recipient_ruc=RE)[0]
        self.assertEqual(items_of(r), [(None, "Servicio Oficina Privada N°716 - Febrero", "1.00", "1080.00")])

    def test_igv_leido_lgv_con_apostrofo(self) -> None:
        from ocr_fiscal.amounts import amount_mentions

        self.assertEqual([(m[0], m[1]) for m in amount_mentions("LG.V (10%): S' 11.36")], [("igv", "11.36")])
        self.assertFalse([m for m in amount_mentions("LGV LOGISTICA 11.36") if m[0] == "igv"])
