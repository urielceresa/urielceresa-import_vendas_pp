import csv
import json
import logging
import os
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk

try:
    import pyautogui
except Exception:  # pragma: no cover - optional dependency
    pyautogui = None

APP_TITLE = "Importador de Vendas"
CONFIG_FILE = "config.json"
LOG_DIR = "logs"
LOG_FILE = os.path.join(LOG_DIR, "app.log")
REQUIRED_FIELDS = ["SRO", "PESO", "ALTURA", "LARGURA", "COMPRIMENTO"]
FIELD_LABELS = {
    "SRO": "SRO",
    "PESO": "Peso",
    "ALTURA": "Altura",
    "LARGURA": "Largura",
    "COMPRIMENTO": "Comprimento",
}


@dataclass
class AppConfig:
    column_mapping: dict
    skip_first_line: bool


def ensure_logging():
    os.makedirs(LOG_DIR, exist_ok=True)
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )


def load_config():
    if not os.path.exists(CONFIG_FILE):
        return AppConfig(column_mapping={}, skip_first_line=True)
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
        return AppConfig(
            column_mapping=data.get("column_mapping", {}),
            skip_first_line=data.get("skip_first_line", True),
        )
    except Exception as exc:
        logging.exception("Falha ao carregar config: %s", exc)
        return AppConfig(column_mapping={}, skip_first_line=True)


def save_config(config: AppConfig):
    with open(CONFIG_FILE, "w", encoding="utf-8") as file:
        json.dump(
            {
                "column_mapping": config.column_mapping,
                "skip_first_line": config.skip_first_line,
            },
            file,
            ensure_ascii=False,
            indent=2,
        )


def detect_delimiter(sample_text):
    try:
        dialect = csv.Sniffer().sniff(sample_text, delimiters=";,\t,")
        return dialect.delimiter
    except csv.Error:
        if ";" in sample_text:
            return ";"
        return ","


def read_csv_rows(file_path, skip_first_line):
    with open(file_path, "r", encoding="utf-8-sig", newline="") as file:
        sample = file.read(4096)
        file.seek(0)
        delimiter = detect_delimiter(sample)
        reader = csv.reader(file, delimiter=delimiter)
        if skip_first_line:
            next(reader, None)
        for row in reader:
            yield row


def count_lines(file_path, skip_first_line):
    return sum(1 for _ in read_csv_rows(file_path, skip_first_line))


def validate_row(data):
    errors = []
    if not data.get("SRO"):
        errors.append("SRO vazio")
    for field in ["PESO", "ALTURA", "LARGURA", "COMPRIMENTO"]:
        value = data.get(field, "")
        try:
            if value == "":
                raise ValueError("vazio")
            float(str(value).replace(",", "."))
        except Exception:
            errors.append(f"{field} inválido")
    return errors


def type_with_enter(value):
    pyautogui.typewrite(str(value))
    pyautogui.press("enter")


def process_file(file_path, column_mapping, skip_first_line, on_progress, should_pause, should_stop):
    total = count_lines(file_path, skip_first_line)
    processed = 0
    logging.info("Iniciando automação para %s", file_path)

    for row in read_csv_rows(file_path, skip_first_line):
        if should_stop.is_set():
            logging.info("Processamento interrompido pelo usuário")
            break
        while should_pause.is_set():
            time.sleep(0.1)
        data = {}
        for field, index in column_mapping.items():
            try:
                data[field] = row[index]
            except Exception:
                data[field] = ""
        errors = validate_row(data)
        if errors:
            logging.warning("Linha %s inválida: %s", processed + 1, "; ".join(errors))
            processed += 1
            on_progress(processed, total)
            continue
        if pyautogui is None:
            logging.error("pyautogui não instalado, não é possível digitar")
            raise RuntimeError("pyautogui não instalado")
        type_with_enter(data["SRO"])
        type_with_enter(data["PESO"])
        type_with_enter(data["ALTURA"])
        type_with_enter(data["LARGURA"])
        type_with_enter(data["COMPRIMENTO"])
        processed += 1
        on_progress(processed, total)

    logging.info("Processamento finalizado")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        ensure_logging()
        self.title(APP_TITLE)
        self.geometry("720x520")
        self.config = load_config()
        self.file_path = ""
        self.total_lines = 0
        self.processing_thread = None
        self.pause_event = threading.Event()
        self.stop_event = threading.Event()
        self._build_ui()

    def _build_ui(self):
        file_frame = ttk.LabelFrame(self, text="Arquivo CSV")
        file_frame.pack(fill="x", padx=10, pady=10)

        self.file_label = ttk.Label(file_frame, text="Nenhum arquivo selecionado")
        self.file_label.pack(side="left", padx=10, pady=10)

        ttk.Button(file_frame, text="Selecionar", command=self.select_file).pack(
            side="right", padx=10, pady=10
        )

        status_frame = ttk.LabelFrame(self, text="Status")
        status_frame.pack(fill="x", padx=10, pady=10)

        self.total_label = ttk.Label(status_frame, text="Total de linhas: 0")
        self.total_label.pack(anchor="w", padx=10)

        self.status_label = ttk.Label(status_frame, text="Status: Aguardando")
        self.status_label.pack(anchor="w", padx=10)

        self.progress = ttk.Progressbar(status_frame, length=400, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=10)

        config_frame = ttk.LabelFrame(self, text="Configuração de Colunas")
        config_frame.pack(fill="both", padx=10, pady=10, expand=True)

        self.skip_var = tk.BooleanVar(value=self.config.skip_first_line)
        ttk.Checkbutton(
            config_frame,
            text="Pular primeira linha",
            variable=self.skip_var,
            command=self.on_skip_toggle,
        ).grid(row=0, column=0, sticky="w", padx=10, pady=5)

        self.column_vars = {}
        for idx, field in enumerate(REQUIRED_FIELDS, start=1):
            ttk.Label(config_frame, text=FIELD_LABELS.get(field, field)).grid(
                row=idx, column=0, sticky="w", padx=10, pady=5
            )
            var = tk.StringVar()
            combo = ttk.Combobox(config_frame, textvariable=var)
            combo.grid(row=idx, column=1, sticky="ew", padx=10, pady=5)
            combo.bind("<<ComboboxSelected>>", lambda _event: self.save_config())
            combo.bind("<KeyRelease>", lambda _event: self.save_config())
            self.column_vars[field] = (var, combo)

        config_frame.columnconfigure(1, weight=1)

        control_frame = ttk.Frame(self)
        control_frame.pack(fill="x", padx=10, pady=10)

        self.start_button = ttk.Button(control_frame, text="Iniciar", command=self.start_processing)
        self.start_button.pack(side="left", padx=5)

        self.pause_button = ttk.Button(control_frame, text="Pausar", command=self.toggle_pause)
        self.pause_button.pack(side="left", padx=5)

        self.stop_button = ttk.Button(control_frame, text="Parar", command=self.stop_processing)
        self.stop_button.pack(side="left", padx=5)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return
        self.file_path = file_path
        self.file_label.config(text=os.path.basename(file_path))
        self.update_columns()
        self.update_total_lines()
        self.status_label.config(text="Status: Aguardando")

    def update_columns(self):
        columns = []
        if self.file_path:
            with open(self.file_path, "r", encoding="utf-8-sig", newline="") as file:
                sample = file.read(4096)
                file.seek(0)
                delimiter = detect_delimiter(sample)
                reader = csv.reader(file, delimiter=delimiter)
                first_row = next(reader, [])
                columns = [col.strip() for col in first_row] or []
        if not columns:
            columns = [str(i) for i in range(1, 51)]

        for field, (var, combo) in self.column_vars.items():
            combo["values"] = columns
            saved = self.config.column_mapping.get(field)
            if saved in columns:
                var.set(saved)
            elif columns:
                var.set(columns[0])

        self.save_config()

    def update_total_lines(self):
        if not self.file_path:
            self.total_lines = 0
            self.total_label.config(text="Total de linhas: 0")
            return
        try:
            self.total_lines = count_lines(self.file_path, self.skip_var.get())
            self.total_label.config(text=f"Total de linhas: {self.total_lines}")
        except Exception as exc:
            logging.exception("Erro ao contar linhas: %s", exc)
            messagebox.showerror("Erro", "Não foi possível ler o arquivo CSV.")

    def save_config(self):
        mapping = {}
        for field, (var, _combo) in self.column_vars.items():
            if var.get():
                mapping[field] = var.get()
        self.config = AppConfig(column_mapping=mapping, skip_first_line=self.skip_var.get())
        save_config(self.config)

    def on_skip_toggle(self):
        self.save_config()
        self.update_total_lines()

    def resolve_mapping(self):
        if not self.file_path:
            return {}
        with open(self.file_path, "r", encoding="utf-8-sig", newline="") as file:
            sample = file.read(4096)
            file.seek(0)
            delimiter = detect_delimiter(sample)
            reader = csv.reader(file, delimiter=delimiter)
            header = next(reader, [])
        mapping = {}
        for field, value in self.config.column_mapping.items():
            if value in header:
                mapping[field] = header.index(value)
            else:
                try:
                    mapping[field] = int(value) - 1
                except Exception:
                    mapping[field] = 0
        return mapping

    def start_processing(self):
        if not self.file_path:
            messagebox.showwarning("Aviso", "Selecione um arquivo CSV primeiro.")
            return
        if self.processing_thread and self.processing_thread.is_alive():
            messagebox.showinfo("Info", "Processamento já está em andamento.")
            return
        if pyautogui is None:
            messagebox.showerror(
                "Erro",
                "pyautogui não está instalado. Instale para habilitar a automação.",
            )
            return
        self.stop_event.clear()
        self.pause_event.clear()
        self.progress["value"] = 0
        self.status_label.config(text="Status: Processando")
        mapping = self.resolve_mapping()
        if not mapping:
            messagebox.showwarning("Aviso", "Configure as colunas antes de iniciar.")
            return
        self.processing_thread = threading.Thread(
            target=self._run_processing,
            args=(mapping,),
            daemon=True,
        )
        self.processing_thread.start()

    def _run_processing(self, mapping):
        try:
            process_file(
                self.file_path,
                mapping,
                self.skip_var.get(),
                self.update_progress,
                self.pause_event,
                self.stop_event,
            )
            self.after(0, lambda: self.status_label.config(text="Status: Finalizado"))
        except Exception as exc:
            logging.exception("Erro durante processamento: %s", exc)
            self.after(0, lambda: messagebox.showerror("Erro", str(exc)))
            self.after(0, lambda: self.status_label.config(text="Status: Erro"))

    def update_progress(self, processed, total):
        self.after(
            0,
            lambda: self._update_progress_ui(processed, total),
        )

    def _update_progress_ui(self, processed, total):
        self.progress["maximum"] = max(total, 1)
        self.progress["value"] = processed
        self.total_label.config(text=f"Total de linhas: {total}")

    def toggle_pause(self):
        if not self.processing_thread or not self.processing_thread.is_alive():
            return
        if self.pause_event.is_set():
            self.pause_event.clear()
            self.pause_button.config(text="Pausar")
            self.status_label.config(text="Status: Processando")
        else:
            self.pause_event.set()
            self.pause_button.config(text="Continuar")
            self.status_label.config(text="Status: Pausado")

    def stop_processing(self):
        if not self.processing_thread or not self.processing_thread.is_alive():
            return
        self.stop_event.set()
        self.pause_event.clear()
        self.pause_button.config(text="Pausar")
        self.status_label.config(text="Status: Parando")


if __name__ == "__main__":
    app = App()
    app.mainloop()
