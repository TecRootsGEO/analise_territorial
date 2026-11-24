#!/usr/bin/env python
"""Interface Tkinter para selecionar arquivos e executar a análise geoespacial."""

from __future__ import annotations

import logging
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from types import SimpleNamespace

from analise_sobreposicao import run


class AnalysisWindow:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Análise de sobreposição geoespacial")
        self.root.minsize(720, 600)
        self.log_queue: queue.Queue[logging.LogRecord] = queue.Queue()
        self.output_path = tk.StringVar()
        self.gpkg_output_path = tk.StringVar()
        self.area_crs = tk.StringVar(value="auto")
        self.source_crs = tk.StringVar(value="EPSG:4674")
        self.chunk_size = tk.StringVar(value="50000")
        self.min_area_ha = tk.StringVar(value="0")
        self.min_percent = tk.StringVar(value="0")
        self.layer_name = tk.StringVar()
        self.glebas_paths: list[Path] = []
        self.operations_paths: list[Path] = []
        self.gpkg_paths: list[Path] = []
        self._build_form()
        self._install_log_handler()
        self.root.after(100, self._drain_logs)
        self.root.protocol("WM_DELETE_WINDOW", self._close)

    def _build_form(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Arquivos de glebas (CSV, GPKG, GeoJSON ou Parquet):").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=4
        )
        self.glebas_list = tk.Listbox(frame, height=3, exportselection=False)
        self.glebas_list.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=(0, 6))
        glebas_buttons = ttk.Frame(frame)
        glebas_buttons.grid(row=1, column=2, sticky="n")
        ttk.Button(glebas_buttons, text="Adicionar...", command=self._add_glebas).pack(
            fill=tk.X, pady=(0, 4)
        )
        ttk.Button(glebas_buttons, text="Remover", command=self._remove_glebas).pack(
            fill=tk.X
        )

        ttk.Label(frame, text="Tabelas auxiliares para junção (CSV ou Parquet, opcional):").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(8, 4)
        )
        self.operations_list = tk.Listbox(frame, height=3, exportselection=False)
        self.operations_list.grid(
            row=3, column=0, columnspan=2, sticky="nsew", padx=(0, 6)
        )
        operation_buttons = ttk.Frame(frame)
        operation_buttons.grid(row=3, column=2, sticky="n")
        ttk.Button(
            operation_buttons, text="Adicionar...", command=self._add_operations
        ).pack(fill=tk.X, pady=(0, 4))
        ttk.Button(
            operation_buttons, text="Remover", command=self._remove_operations
        ).pack(fill=tk.X)

        ttk.Label(frame, text="Camada Base (GeoPackage, GeoJSON ou GeoParquet):").grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(12, 4)
        )
        self.gpkg_list = tk.Listbox(frame, height=4, exportselection=False)
        self.gpkg_list.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=(0, 6))
        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=2, sticky="n")
        ttk.Button(buttons, text="Adicionar...", command=self._add_geopackages).pack(
            fill=tk.X, pady=(0, 4)
        )
        ttk.Button(buttons, text="Remover", command=self._remove_geopackage).pack(
            fill=tk.X
        )

        ttk.Label(frame, text="Camada (opcional):").grid(row=6, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.layer_name).grid(
            row=6, column=1, sticky="ew", padx=6
        )

        ttk.Label(frame, text="Arquivo Excel de saída:").grid(
            row=7, column=0, sticky="w", pady=4
        )
        ttk.Entry(frame, textvariable=self.output_path).grid(
            row=7, column=1, sticky="ew", padx=6
        )
        ttk.Button(frame, text="Salvar como...", command=self._browse_output).grid(
            row=7, column=2
        )

        ttk.Label(frame, text="GeoPackage das glebas sobrepostas (opcional):").grid(
            row=8, column=0, sticky="w", pady=4
        )
        ttk.Entry(frame, textvariable=self.gpkg_output_path).grid(
            row=8, column=1, sticky="ew", padx=6
        )
        ttk.Button(frame, text="Salvar como...", command=self._browse_gpkg_output).grid(
            row=8, column=2
        )

        options = ttk.LabelFrame(frame, text="Opções", padding=8)
        options.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        for column in range(4):
            options.columnconfigure(column, weight=1)
        ttk.Label(options, text="CRS das glebas:").grid(row=0, column=0, sticky="w")
        ttk.Entry(options, textvariable=self.source_crs, width=16).grid(
            row=0, column=1, sticky="w", padx=(4, 12)
        )
        ttk.Label(options, text="CRS para áreas (auto ou EPSG):").grid(
            row=0, column=2, sticky="w"
        )
        ttk.Entry(options, textvariable=self.area_crs, width=16).grid(
            row=0, column=3, sticky="w", padx=4
        )
        ttk.Label(options, text="Linhas por bloco:").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(options, textvariable=self.chunk_size, width=16).grid(
            row=1, column=1, sticky="w", padx=(4, 12), pady=6
        )
        ttk.Label(options, text="Área mín. (ha):").grid(row=1, column=2, sticky="w", pady=6)
        ttk.Entry(options, textvariable=self.min_area_ha, width=16).grid(
            row=1, column=3, sticky="w", padx=4, pady=6
        )
        ttk.Label(options, text="Percentual mín. (%):").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(options, textvariable=self.min_percent, width=16).grid(
            row=2, column=1, sticky="w", padx=(4, 12), pady=2
        )

        self.run_button = ttk.Button(frame, text="Iniciar análise", command=self._start)
        self.run_button.grid(row=10, column=0, columnspan=3, pady=8)
        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.grid(row=11, column=0, columnspan=3, sticky="ew")

        ttk.Label(frame, text="Progresso:").grid(row=12, column=0, sticky="w", pady=(12, 4))
        self.log_view = tk.Text(frame, height=10, state=tk.DISABLED, wrap=tk.WORD)
        self.log_view.grid(row=13, column=0, columnspan=2, sticky="nsew")
        log_scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.log_view.yview)
        log_scrollbar.grid(row=13, column=2, sticky="ns")
        self.log_view.configure(yscrollcommand=log_scrollbar.set)
        frame.rowconfigure(13, weight=1)

    def _install_log_handler(self) -> None:
        class QueueLogHandler(logging.Handler):
            def __init__(self, destination: queue.Queue[logging.LogRecord]) -> None:
                super().__init__()
                self.destination = destination

            def emit(self, record: logging.LogRecord) -> None:
                self.destination.put(record)

        self.log_handler = QueueLogHandler(self.log_queue)
        self.log_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        root_logger = logging.getLogger()
        root_logger.addHandler(self.log_handler)
        root_logger.setLevel(logging.INFO)

    def _add_glebas(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Selecionar arquivos de glebas",
            filetypes=[
                ("Arquivos suportados", "*.csv *.gpkg *.geojson *.json *.parquet"),
                ("Arquivos CSV", "*.csv"),
                ("GeoPackage / GeoJSON", "*.gpkg *.geojson *.json"),
                ("Parquet / GeoParquet", "*.parquet"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        for raw_path in paths:
            path = Path(raw_path)
            if path not in self.glebas_paths:
                self.glebas_paths.append(path)
                self.glebas_list.insert(tk.END, str(path))

    def _remove_glebas(self) -> None:
        selected = self.glebas_list.curselection()
        for index in reversed(selected):
            self.glebas_list.delete(index)
            del self.glebas_paths[index]

    def _add_operations(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Selecionar tabelas auxiliares para junção",
            filetypes=[
                ("Tabelas suportadas", "*.csv *.parquet"),
                ("Arquivos CSV", "*.csv"),
                ("Parquet", "*.parquet"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        for raw_path in paths:
            path = Path(raw_path)
            if path not in self.operations_paths:
                self.operations_paths.append(path)
                self.operations_list.insert(tk.END, str(path))

    def _remove_operations(self) -> None:
        selected = self.operations_list.curselection()
        for index in reversed(selected):
            self.operations_list.delete(index)
            del self.operations_paths[index]

    def _add_geopackages(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Selecionar camadas da Camada Base",
            filetypes=[
                ("Camadas espaciais", "*.gpkg *.geojson *.json *.parquet"),
                ("GeoPackage", "*.gpkg"),
                ("GeoJSON", "*.geojson *.json"),
                ("GeoParquet", "*.parquet"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        for raw_path in paths:
            path = Path(raw_path)
            if path not in self.gpkg_paths:
                self.gpkg_paths.append(path)
                self.gpkg_list.insert(tk.END, str(path))

    def _remove_geopackage(self) -> None:
        selected = self.gpkg_list.curselection()
        for index in reversed(selected):
            self.gpkg_list.delete(index)
            del self.gpkg_paths[index]

    def _browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Salvar resultado",
            defaultextension=".xlsx",
            filetypes=[("Pasta de trabalho Excel", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    def _browse_gpkg_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Salvar GeoPackage das glebas sobrepostas",
            defaultextension=".gpkg",
            filetypes=[("GeoPackage", "*.gpkg")],
        )
        if path:
            self.gpkg_output_path.set(path)

    def _start(self) -> None:
        if not self.glebas_paths:
            messagebox.showerror(
                "Dados incompletos",
                "Selecione ao menos um arquivo de glebas.",
            )
            return
        if not self.gpkg_paths:
            messagebox.showerror(
                "Dados incompletos", "Adicione ao menos um arquivo da Camada Base."
            )
            return
        if not self.output_path.get():
            messagebox.showerror("Dados incompletos", "Selecione o arquivo Excel de saída.")
            return
        try:
            chunk_size = int(self.chunk_size.get())
            if chunk_size < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Opção inválida", "O tamanho do bloco deve ser um inteiro positivo.")
            return
        try:
            min_area_ha = float((self.min_area_ha.get() or "0").strip().replace(",", "."))
            if min_area_ha < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Opção inválida", "A área mínima de sobreposição deve ser um número maior ou igual a zero."
            )
            return
        try:
            min_percent = float((self.min_percent.get() or "0").strip().replace(",", "."))
            if min_percent < 0 or min_percent > 100:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Opção inválida", "O percentual mínimo de sobreposição deve estar entre 0 e 100."
            )
            return

        args = SimpleNamespace(
            glebas=list(self.glebas_paths),
            operacoes=list(self.operations_paths),
            terras_indigenas=list(self.gpkg_paths),
            saida=Path(self.output_path.get()),
            saida_gpkg=(
                Path(self.gpkg_output_path.get()) if self.gpkg_output_path.get() else None
            ),
            camada_ti=self.layer_name.get().strip() or None,
            coluna_ref=None,
            coluna_geometria=None,
            coluna_nome_ti=None,
            coluna_valor=None,
            crs_glebas=self.source_crs.get().strip(),
            crs_area=self.area_crs.get().strip(),
            area_minima_ha=min_area_ha,
            percentual_minimo=min_percent,
            codificacao="auto",
            delimitador=None,
            tamanho_bloco=chunk_size,
            verbose=False,
        )
        self.run_button.configure(state=tk.DISABLED)
        self.progress.start(12)
        threading.Thread(target=self._run_analysis, args=(args,), daemon=True).start()

    def _run_analysis(self, args: SimpleNamespace) -> None:
        try:
            run(args)
            self.log_queue.put(
                logging.LogRecord(
                    "analise_sobreposicao",
                    logging.INFO,
                    "",
                    0,
                    "Análise concluída com sucesso.",
                    (),
                    None,
                )
            )
            saved_paths = [str(args.saida)]
            if args.saida_gpkg is not None:
                saved_paths.append(str(args.saida_gpkg))
            success_message = "Arquivos salvos em:\n" + "\n".join(saved_paths)
            self.root.after(
                0,
                lambda message=success_message: messagebox.showinfo(
                    "Concluído", message
                ),
            )
        except Exception as error:
            logging.getLogger("analise_sobreposicao").exception(
                "Falha na análise: %s", error
            )
            self.root.after(
                0,
                lambda error_message=str(error): messagebox.showerror(
                    "Falha na análise", error_message
                ),
            )
        finally:
            self.root.after(0, self._finish)

    def _finish(self) -> None:
        self.progress.stop()
        self.run_button.configure(state=tk.NORMAL)

    def _drain_logs(self) -> None:
        self.log_view.configure(state=tk.NORMAL)
        while True:
            try:
                record = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.log_view.insert(tk.END, self.log_handler.format(record) + "\n")
        self.log_view.see(tk.END)
        self.log_view.configure(state=tk.DISABLED)
        self.root.after(100, self._drain_logs)

    def _close(self) -> None:
        logging.getLogger().removeHandler(self.log_handler)
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    AnalysisWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
