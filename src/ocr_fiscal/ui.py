from __future__ import annotations

from dataclasses import replace
import os
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any

from .config import AppConfig
from .exports import OUTPUT_FILENAME, export_results
from .processor import FiscalProcessor, ProgressUpdate, ScanSummary
from .tax import validate_ruc


# La ventana del programa (Tkinter): se eligen las carpetas y el RUC, se inicia o se cancela el análisis,
# se ve el avance y se abre el Excel.
class CompactFiscalUI:
    """Ventana para elegir las carpetas y seguir la lectura de los comprobantes."""

    # Arma la ventana y las variables que muestran el estado. El análisis corre en otro hilo y se comunica
    # con la ventana a través de una cola, para que la pantalla no se congele mientras se leen los documentos.
    def __init__(
        self,
        root: tk.Tk,
        *,
        initial_input: str | Path | None = None,
        initial_output: str | Path | None = None,
        config: AppConfig | None = None,
    ) -> None:
        self.root = root
        self.config = config or AppConfig()
        self.root.title("OCR CDP - Detección de Comprobantes y Contenedores ZIP")
        self.root.geometry("760x480")
        self.root.minsize(680, 420)
        self.root.resizable(True, True)

        # Variables ligadas a los campos de la pantalla: carpetas, RUC y casilla de los ZIP.
        self.input_var = tk.StringVar(value=str(initial_input or Path.cwd() / "entrada"))
        self.output_var = tk.StringVar(value=str(initial_output or Path.cwd() / "salida"))
        self.ruc_var = tk.StringVar()
        self.inspect_zip_var = tk.BooleanVar(value=self.config.inspect_archives)

        # Variables de los indicadores de avance y de los contadores.
        self.stage_var = tk.StringVar(value="Listo para iniciar")
        self.current_file_var = tk.StringVar(value="-")
        self.total_var = tk.StringVar(value="0 / 0")
        self.zip_var = tk.StringVar(value="0")
        self.detected_var = tk.StringVar(value="0")
        self.ambiguous_var = tk.StringVar(value="0")
        self.non_receipt_var = tk.StringVar(value="0")
        self.errors_var = tk.StringVar(value="0")
        self.progress_value = tk.DoubleVar(value=0.0)

        # La cola por donde el hilo del análisis manda mensajes a la ventana, y la bandera para cancelar.
        self._queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._cancel_event = threading.Event()
        self._worker: threading.Thread | None = None
        self._last_output_root: Path | None = None

        # Dibujamos la ventana y empezamos a revisar la cola de mensajes cada 100 milisegundos.
        self._build_layout()
        self.root.after(100, self._poll_queue)

    # Dibuja los cinco bloques de la ventana: datos de la corrida, botones, avance, contadores y bitácora.
    def _build_layout(self) -> None:
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)

        # 1. Carpetas de entrada y de salida, RUC del comprador y casilla de los ZIP.
        paths_frame = ttk.LabelFrame(main_frame, text="Configuración de corrida", padding=8)
        paths_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 6))
        paths_frame.columnconfigure(1, weight=1)

        ttk.Label(paths_frame, text="Carpeta de entrada:").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        self.input_entry = ttk.Entry(paths_frame, textvariable=self.input_var)
        self.input_entry.grid(row=0, column=1, sticky="ew", pady=2)
        self.browse_in_btn = ttk.Button(paths_frame, text="Examinar...", command=self._browse_input)
        self.browse_in_btn.grid(row=0, column=2, padx=(6, 0), pady=2)

        ttk.Label(paths_frame, text="Carpeta de salida:").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=2)
        self.output_entry = ttk.Entry(paths_frame, textvariable=self.output_var)
        self.output_entry.grid(row=1, column=1, sticky="ew", pady=2)
        self.browse_out_btn = ttk.Button(paths_frame, text="Examinar...", command=self._browse_output)
        self.browse_out_btn.grid(row=1, column=2, padx=(6, 0), pady=2)

        opts_frame = ttk.Frame(paths_frame)
        opts_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(4, 0))
        ttk.Label(opts_frame, text="RUC receptor objetivo:").pack(side=tk.LEFT, padx=(0, 4))
        self.ruc_entry = ttk.Entry(opts_frame, textvariable=self.ruc_var, width=15)
        self.ruc_entry.pack(side=tk.LEFT, padx=(0, 14))
        self.zip_chk = ttk.Checkbutton(
            opts_frame,
            text="Inspeccionar ZIP y ZIP anidados",
            variable=self.inspect_zip_var,
        )
        self.zip_chk.pack(side=tk.LEFT, padx=(0, 14))

        # 2. Botones: iniciar, cancelar y abrir los resultados.
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 6))
        self.start_btn = ttk.Button(btn_frame, text="Iniciar análisis", command=self.start_analysis)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.cancel_btn = ttk.Button(btn_frame, text="Cancelar", command=self.cancel_analysis, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=(0, 16))
        self.open_folder_btn = ttk.Button(btn_frame, text="Abrir carpeta de salida", command=self.open_output_folder)
        self.open_folder_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.open_excel_btn = ttk.Button(btn_frame, text=f"Abrir {OUTPUT_FILENAME}", command=self.open_excel_file)
        self.open_excel_btn.pack(side=tk.LEFT)

        # 3. Barra de avance y archivo que se está leyendo.
        prog_frame = ttk.LabelFrame(main_frame, text="Estado y avance", padding=8)
        prog_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 6))
        prog_frame.columnconfigure(1, weight=1)

        self.progressbar = ttk.Progressbar(
            prog_frame, variable=self.progress_value, maximum=100.0, mode="determinate"
        )
        self.progressbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))

        ttk.Label(prog_frame, text="Etapa:").grid(row=1, column=0, sticky="w", padx=(0, 6))
        ttk.Label(prog_frame, textvariable=self.stage_var).grid(row=1, column=1, sticky="w")
        ttk.Label(prog_frame, text="Actual:").grid(row=2, column=0, sticky="w", padx=(0, 6))
        ttk.Label(prog_frame, textvariable=self.current_file_var, width=78).grid(row=2, column=1, sticky="w")

        # 4. Contadores en vivo: unidades, ZIP, comprobantes, ambiguos, no comprobantes y errores.
        counts_frame = ttk.Frame(main_frame)
        counts_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 6))
        counter_specs = [
            ("Unidades:", self.total_var),
            ("ZIPs:", self.zip_var),
            ("Comprobantes:", self.detected_var),
            ("Ambiguos:", self.ambiguous_var),
            ("No comprobantes:", self.non_receipt_var),
            ("Errores:", self.errors_var),
        ]
        for idx, (label_text, var) in enumerate(counter_specs):
            counts_frame.columnconfigure(idx * 2 + 1, weight=1)
            ttk.Label(counts_frame, text=label_text).grid(row=0, column=idx * 2, sticky="w", padx=(0, 2))
            ttk.Label(counts_frame, textvariable=var, font=("Segoe UI", 9, "bold")).grid(
                row=0, column=idx * 2 + 1, sticky="w", padx=(0, 10)
            )

        # 5. Bitácora con los últimos eventos.
        log_frame = ttk.LabelFrame(main_frame, text="Bitácora de eventos recientes", padding=6)
        log_frame.grid(row=4, column=0, columnspan=3, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=6, wrap="word", state=tk.DISABLED, font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    # Añade un mensaje a la bitácora y la limita a las últimas 200 líneas.
    def _append_log(self, message: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message.rstrip() + "\n")
        line_count = int(self.log_text.index("end-1c").split(".")[0])
        if line_count > 200:
            self.log_text.delete("1.0", f"{line_count - 200}.0")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    # Abre el selector de carpeta para la entrada.
    def _browse_input(self) -> None:
        selected = filedialog.askdirectory(
            parent=self.root, title="Elige la carpeta de comprobantes",
            initialdir=self.input_var.get() or str(Path.cwd()), mustexist=True,
        )
        if selected:
            self.input_var.set(selected)

    # Abre el selector de carpeta para la salida (puede no existir todavía).
    def _browse_output(self) -> None:
        selected = filedialog.askdirectory(
            parent=self.root, title=f"Elige dónde guardar {OUTPUT_FILENAME}",
            initialdir=self.output_var.get() or str(Path.cwd()), mustexist=False,
        )
        if selected:
            self.output_var.set(selected)

    # Comprueba lo que escribió el usuario antes de empezar: las carpetas, que la salida no esté dentro de la entrada
    # y que el RUC sea válido. Si algo falla, lanza un ValueError con el motivo.
    def validate_inputs(self) -> tuple[Path, Path, str]:
        input_folder = self.input_var.get().strip()
        output_folder = self.output_var.get().strip()
        if not input_folder or not output_folder:
            raise ValueError("Elige las carpetas de entrada y salida.")
        input_path = Path(input_folder).resolve()
        output_path = Path(output_folder).resolve()
        ruc = self.ruc_var.get().strip()

        if not input_path.is_dir():
            raise ValueError(f"La carpeta de entrada no existe: {input_path}")
        if output_path == input_path or input_path in output_path.parents:
            raise ValueError("Elija una carpeta de salida fuera de la entrada.")

        val = validate_ruc(ruc)
        if not (val.format_valid and val.checksum_valid):
            raise ValueError(f"El RUC receptor objetivo '{ruc}' no tiene formato o dígito verificador válido")

        return input_path, output_path, ruc

    # Botón «Iniciar análisis»: valida, deja la pantalla lista y lanza el análisis en un hilo aparte.
    def start_analysis(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            return
        try:
            input_path, output_path, ruc = self.validate_inputs()
        except ValueError as exc:
            messagebox.showerror("Validación de parámetros", str(exc), parent=self.root)
            return

        # Reiniciamos el avance y bloqueamos el botón de iniciar mientras corre el análisis.
        self._cancel_event.clear()
        self._last_output_root = output_path
        self.progress_value.set(0.0)
        self.stage_var.set("Iniciando...")
        self.current_file_var.set(str(input_path))
        self.start_btn.configure(state=tk.DISABLED)
        self.cancel_btn.configure(state=tk.NORMAL)
        self._append_log(f"Iniciando análisis en: {input_path}")
        self._append_log(f"Salida configurada en: {output_path} | RUC receptor: {ruc}")

        inspect_zips = bool(self.inspect_zip_var.get())

        # El análisis se hace en un hilo aparte para no congelar la ventana.
        self._worker = threading.Thread(
            target=self._run_worker,
            args=(input_path, output_path, ruc, inspect_zips),
            daemon=True,
        )
        self._worker.start()

    # Botón «Cancelar»: avisa al análisis, que se detiene al terminar el documento que está leyendo.
    def cancel_analysis(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            self._cancel_event.set()
            self.stage_var.set("Cancelando al terminar unidad actual...")
            self._append_log("Solicitud de cancelación enviada; esperando cierre seguro de la unidad en curso...")

    # Lo que hace el hilo: arma los ajustes con lo elegido en pantalla, corre el análisis, genera el Excel
    # y avisa a la ventana a través de la cola.
    def _run_worker(
        self,
        input_path: Path,
        output_path: Path,
        ruc: str,
        inspect_zips: bool,
    ) -> None:
        try:
            # Se parte de los ajustes del config.json y se pisan el RUC y la casilla de ZIP con lo elegido en pantalla.
            config = replace(
                self.config,
                recipient_ruc=ruc,
                inspect_archives=inspect_zips,
                inspect_nested_archives=inspect_zips,
            )
            processor = FiscalProcessor(output_path, config)

            # Cada aviso de avance se pone en la cola; la ventana lo recoge enseguida.
            def on_progress(update: ProgressUpdate) -> None:
                self._queue.put(("progress", update))

            summary = processor.scan(
                input_path,
                progress_callback=on_progress,
                cancel_event=self._cancel_event,
            )
            workbook = export_results(processor.store, output_path, summary=summary)
            self._queue.put(("done", (summary, workbook)))
        except Exception as exc:
            self._queue.put(("error", f"{type(exc).__name__}: {exc}"))

    # Cada 100 ms vacía la cola de mensajes y actualiza la pantalla.
    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                self._handle_queue_item(kind, payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # Reacciona a un mensaje del hilo: «progress» (avance), «done» (terminó) o «error».
    def _handle_queue_item(self, kind: str, payload: Any) -> None:
        # Avance: se actualizan la barra, la etapa y los contadores.
        if kind == "progress":
            upd: ProgressUpdate = payload
            pct = (upd.current_unit / upd.total_units * 100.0) if upd.total_units > 0 else 0.0
            self.progress_value.set(round(pct, 1))
            self.stage_var.set(f"{upd.stage.capitalize()} ({upd.current_unit}/{upd.total_units})")
            self.current_file_var.set(upd.current_path)
            self.total_var.set(f"{upd.current_unit} / {upd.total_units}")
            self.zip_var.set(str(upd.zip_container_count))
            self.detected_var.set(str(upd.detected_count))
            self.ambiguous_var.set(str(upd.ambiguous_count))
            self.non_receipt_var.set(str(upd.no_comprobante_count))
            self.errors_var.set(str(upd.read_error_count))
            if upd.message:
                self._append_log(f"[{upd.current_unit}/{upd.total_units}] {upd.message}")
        # Terminó: se muestra el total (o el resultado parcial, si se canceló) y se vuelven a habilitar los botones.
        elif kind == "done":
            summary, workbook = payload
            assert isinstance(summary, ScanSummary)
            denom = summary.discovered_unit_count if summary.discovered_unit_count > 0 else summary.processed_file_count
            pct = (summary.processed_file_count / denom * 100.0) if denom > 0 else (100.0 if not summary.cancelled else 0.0)
            self.progress_value.set(100.0 if not summary.cancelled else round(pct, 1))
            status_label = (
                f"Cancelado (resultados parciales: {summary.processed_file_count}/{denom} procesadas, {summary.pending_unit_count} pendientes)"
                if summary.cancelled
                else "Completado"
            )
            self.stage_var.set(status_label)
            self.total_var.set(f"{summary.processed_file_count} / {denom}")
            self.zip_var.set(str(summary.zip_container_count))
            self.detected_var.set(str(summary.detected_count))
            self.ambiguous_var.set(str(summary.ambiguous_count))
            self.non_receipt_var.set(str(summary.no_comprobante_count))
            self.errors_var.set(str(summary.read_error_count))
            self._append_log(
                f"{status_label}: {summary.detected_count} CDP, "
                f"{summary.ambiguous_count} indeterminados, {summary.no_comprobante_count} no CDP, "
                f"{summary.read_error_count} errores de lectura. Excel en {workbook}"
            )
            self.start_btn.configure(state=tk.NORMAL)
            self.cancel_btn.configure(state=tk.DISABLED)
        # Falló: se muestra el error en la bitácora.
        elif kind == "error":
            self.stage_var.set("Error en ejecución")
            self._append_log(f"ERROR: {payload}")
            self.start_btn.configure(state=tk.NORMAL)
            self.cancel_btn.configure(state=tk.DISABLED)

    # Abre la carpeta de salida en el explorador (sólo funciona en Windows).
    def open_output_folder(self) -> None:
        target = self._last_output_root or Path(self.output_var.get().strip())
        if target.exists() and hasattr(os, "startfile"):
            os.startfile(str(target))

    # Abre CDP.xlsx con el programa predeterminado (sólo funciona en Windows).
    def open_excel_file(self) -> None:
        target = (self._last_output_root or Path(self.output_var.get().strip())) / OUTPUT_FILENAME
        if target.is_file() and hasattr(os, "startfile"):
            os.startfile(str(target))


# Abre la ventana y espera a que el usuario la cierre. Si no hay entorno gráfico, explica cómo usar la terminal.
def launch_ui(
    *,
    input_root: str | Path | None = None,
    output_root: str | Path | None = None,
    config: AppConfig | None = None,
) -> int:
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        raise ValueError("No se pudo abrir una ventana. Puedes ejecutar iniciar.py scan desde la terminal.") from exc
    CompactFiscalUI(
        root,
        initial_input=input_root,
        initial_output=output_root,
        config=config,
    )
    root.mainloop()
    return 0
