 (cd "$(git rev-parse --show-toplevel)" && git apply --3way <<'EOF' 
diff --git a/main.py b/main.py
new file mode 100644
index 0000000000000000000000000000000000000000..02399bc03970a5cba7eab75280320d0bb5bd1668
--- /dev/null
+++ b/main.py
@@ -0,0 +1,685 @@
+import csv
+import json
+import logging
+import os
+import threading
+import time
+import tkinter as tk
+from dataclasses import dataclass
+from tkinter import filedialog, messagebox, ttk
+
+try:
+    import pyautogui
+except Exception:  # pragma: no cover - optional dependency
+    pyautogui = None
+
+APP_TITLE = "Importador de Vendas"
+CONFIG_FILE = "config.json"
+LOG_DIR = "logs"
+LOG_FILE = os.path.join(LOG_DIR, "app.log")
+REQUIRED_FIELDS = ["OBJETO", "PESO", "ALTURA", "LARGURA", "COMPRIMENTO"]
+FIELD_LABELS = {
+    "OBJETO": "Objeto",
+    "PESO": "Peso",
+    "ALTURA": "Altura",
+    "LARGURA": "Largura",
+    "COMPRIMENTO": "Comprimento",
+}
+
+
+@dataclass
+class AppConfig:
+    column_mapping: dict
+    skip_first_line: bool
+    speed_preset: str
+    speed_slider: int
+    auto_speed: bool
+
+
+def ensure_logging():
+    os.makedirs(LOG_DIR, exist_ok=True)
+    logging.basicConfig(
+        filename=LOG_FILE,
+        level=logging.INFO,
+        format="%(asctime)s [%(levelname)s] %(message)s",
+    )
+
+
+def load_config():
+    if not os.path.exists(CONFIG_FILE):
+        return AppConfig(
+            column_mapping={},
+            skip_first_line=True,
+            speed_preset="Normal",
+            speed_slider=50,
+            auto_speed=False,
+        )
+    try:
+        with open(CONFIG_FILE, "r", encoding="utf-8") as file:
+            data = json.load(file)
+        column_mapping = data.get("column_mapping", {})
+        if "SRO" in column_mapping and "OBJETO" not in column_mapping:
+            column_mapping["OBJETO"] = column_mapping.pop("SRO")
+        return AppConfig(
+            column_mapping=column_mapping,
+            skip_first_line=data.get("skip_first_line", True),
+            speed_preset=data.get("speed_preset", "Normal"),
+            speed_slider=int(data.get("speed_slider", 50)),
+            auto_speed=bool(data.get("auto_speed", False)),
+        )
+    except Exception as exc:
+        logging.exception("Falha ao carregar config: %s", exc)
+        return AppConfig(
+            column_mapping={},
+            skip_first_line=True,
+            speed_preset="Normal",
+            speed_slider=50,
+            auto_speed=False,
+        )
+
+
+def save_config(config: AppConfig):
+    with open(CONFIG_FILE, "w", encoding="utf-8") as file:
+        json.dump(
+            {
+                "column_mapping": config.column_mapping,
+                "skip_first_line": config.skip_first_line,
+                "speed_preset": config.speed_preset,
+                "speed_slider": config.speed_slider,
+                "auto_speed": config.auto_speed,
+            },
+            file,
+            ensure_ascii=False,
+            indent=2,
+        )
+
+
+def detect_delimiter(sample_text):
+    try:
+        dialect = csv.Sniffer().sniff(sample_text, delimiters=";,\t,")
+        return dialect.delimiter
+    except csv.Error:
+        if ";" in sample_text:
+            return ";"
+        return ","
+
+
+def read_csv_rows(file_path, skip_first_line):
+    with open(file_path, "r", encoding="utf-8-sig", newline="") as file:
+        sample = file.read(4096)
+        file.seek(0)
+        delimiter = detect_delimiter(sample)
+        reader = csv.reader(file, delimiter=delimiter)
+        if skip_first_line:
+            next(reader, None)
+        for row in reader:
+            yield row
+
+
+def count_lines(file_path, skip_first_line):
+    return sum(1 for _ in read_csv_rows(file_path, skip_first_line))
+
+
+def detect_header(file_path):
+    with open(file_path, "r", encoding="utf-8-sig", newline="") as file:
+        sample = file.read(4096)
+        delimiter = detect_delimiter(sample)
+        try:
+            return csv.Sniffer().has_header(sample), delimiter
+        except csv.Error:
+            return False, delimiter
+
+
+def normalize_text(value):
+    return "".join(ch.lower() for ch in str(value).strip() if ch.isalnum())
+
+
+def validate_row(data):
+    errors = []
+    if not data.get("OBJETO"):
+        errors.append("Objeto vazio")
+    for field in ["PESO", "ALTURA", "LARGURA", "COMPRIMENTO"]:
+        value = data.get(field, "")
+        try:
+            if value == "":
+                raise ValueError("vazio")
+            float(str(value).replace(",", "."))
+        except Exception:
+            errors.append(f"{field} inválido")
+    return errors
+
+
+def type_with_enter(value, key_interval):
+    pyautogui.typewrite(str(value), interval=key_interval)
+    pyautogui.press("enter")
+
+
+def process_file(
+    file_path,
+    column_mapping,
+    skip_first_line,
+    on_progress,
+    should_pause,
+    should_stop,
+    speed_settings,
+):
+    total = count_lines(file_path, skip_first_line)
+    processed = 0
+    auto_multiplier = 1.0
+    logging.info("Iniciando automação para %s", file_path)
+
+    for row in read_csv_rows(file_path, skip_first_line):
+        if should_stop.is_set():
+            logging.info("Processamento interrompido pelo usuário")
+            break
+        while should_pause.is_set():
+            time.sleep(0.1)
+        data = {}
+        for field, index in column_mapping.items():
+            try:
+                data[field] = row[index]
+            except Exception:
+                data[field] = ""
+        errors = validate_row(data)
+        if errors:
+            logging.warning("Linha %s inválida: %s", processed + 1, "; ".join(errors))
+            if speed_settings["auto_speed"]:
+                auto_multiplier = min(auto_multiplier + 0.1, 3.0)
+            processed += 1
+            on_progress(processed, total)
+            continue
+        if pyautogui is None:
+            logging.error("pyautogui não instalado, não é possível digitar")
+            raise RuntimeError("pyautogui não instalado")
+        key_interval = speed_settings["key_interval"] * auto_multiplier
+        field_delay = speed_settings["field_delay"] * auto_multiplier
+        type_with_enter(data["OBJETO"], key_interval)
+        time.sleep(field_delay)
+        type_with_enter(data["PESO"], key_interval)
+        time.sleep(field_delay)
+        type_with_enter(data["ALTURA"], key_interval)
+        time.sleep(field_delay)
+        type_with_enter(data["LARGURA"], key_interval)
+        time.sleep(field_delay)
+        type_with_enter(data["COMPRIMENTO"], key_interval)
+        time.sleep(field_delay)
+        if speed_settings["auto_speed"] and auto_multiplier > 1.0:
+            auto_multiplier = max(auto_multiplier - 0.05, 1.0)
+        processed += 1
+        on_progress(processed, total)
+
+    logging.info("Processamento finalizado")
+
+
+class App(tk.Tk):
+    def __init__(self):
+        super().__init__()
+        ensure_logging()
+        self.title(APP_TITLE)
+        self.geometry("360x520")
+        self.config = load_config()
+        self.file_path = ""
+        self.total_lines = 0
+        self.processing_thread = None
+        self.pause_event = threading.Event()
+        self.stop_event = threading.Event()
+        self.countdown_seconds = 5
+        self.countdown_remaining = 0
+        self.countdown_after_id = None
+        self.column_values = []
+        self.column_search_vars = {}
+        self.has_header = True
+        self.prep_window = None
+        self.speed_preset_var = tk.StringVar(value=self.config.speed_preset)
+        self.speed_slider_var = tk.IntVar(value=self.config.speed_slider)
+        self.auto_speed_var = tk.BooleanVar(value=self.config.auto_speed)
+        self._build_ui()
+        self.bind("<Escape>", lambda _event: self.stop_processing())
+
+    def _build_ui(self):
+        file_frame = ttk.LabelFrame(self, text="Arquivo CSV")
+        file_frame.pack(fill="x", padx=10, pady=10)
+
+        self.file_label = ttk.Label(file_frame, text="Nenhum arquivo selecionado")
+        self.file_label.pack(side="left", padx=10, pady=10)
+
+        ttk.Button(file_frame, text="Selecionar", command=self.select_file).pack(
+            side="right", padx=10, pady=10
+        )
+
+        status_frame = ttk.LabelFrame(self, text="Status")
+        status_frame.pack(fill="x", padx=10, pady=10)
+
+        self.total_label = ttk.Label(status_frame, text="Total de linhas: 0")
+        self.total_label.pack(anchor="w", padx=10)
+
+        self.status_label = ttk.Label(status_frame, text="Status: Aguardando")
+        self.status_label.pack(anchor="w", padx=10)
+
+        self.countdown_label = ttk.Label(status_frame, text="Início em: -")
+        self.countdown_label.pack(anchor="w", padx=10)
+
+        self.progress = ttk.Progressbar(status_frame, length=400, mode="determinate")
+        self.progress.pack(fill="x", padx=10, pady=10)
+
+        config_frame = ttk.LabelFrame(self, text="Configuração de Colunas")
+        config_frame.pack(fill="both", padx=10, pady=10, expand=True)
+
+        self.skip_var = tk.BooleanVar(value=self.config.skip_first_line)
+        ttk.Checkbutton(
+            config_frame,
+            text="Pular primeira linha",
+            variable=self.skip_var,
+            command=self.on_skip_toggle,
+        ).grid(row=0, column=0, sticky="w", padx=10, pady=5)
+
+        self.column_vars = {}
+        for idx, field in enumerate(REQUIRED_FIELDS, start=1):
+            ttk.Label(config_frame, text=FIELD_LABELS.get(field, field)).grid(
+                row=idx, column=0, sticky="w", padx=10, pady=5
+            )
+            search_var = tk.StringVar()
+            search_entry = ttk.Entry(config_frame, textvariable=search_var)
+            search_entry.grid(row=idx, column=1, sticky="ew", padx=10, pady=5)
+            search_entry.bind(
+                "<KeyRelease>",
+                lambda event, field=field: self.on_column_search(event, field),
+            )
+
+            var = tk.StringVar()
+            combo = ttk.Combobox(config_frame, textvariable=var, state="readonly")
+            combo.grid(row=idx, column=2, sticky="ew", padx=10, pady=5)
+            combo.bind("<<ComboboxSelected>>", lambda _event: self.save_config())
+
+            self.column_search_vars[field] = (search_var, search_entry)
+            self.column_vars[field] = (var, combo)
+
+        config_frame.columnconfigure(1, weight=1)
+        config_frame.columnconfigure(2, weight=1)
+
+        ttk.Button(config_frame, text="Salvar configurações", command=self.save_config).grid(
+            row=len(REQUIRED_FIELDS) + 1,
+            column=0,
+            columnspan=3,
+            sticky="ew",
+            padx=10,
+            pady=10,
+        )
+
+        speed_frame = ttk.LabelFrame(self, text="Velocidade da automação")
+        speed_frame.pack(fill="x", padx=10, pady=10)
+
+        ttk.Label(speed_frame, text="Preset").grid(row=0, column=0, sticky="w", padx=10, pady=5)
+        speed_combo = ttk.Combobox(
+            speed_frame,
+            textvariable=self.speed_preset_var,
+            values=["Lenta", "Normal", "Rápida"],
+            state="readonly",
+        )
+        speed_combo.grid(row=0, column=1, sticky="ew", padx=10, pady=5)
+        speed_combo.bind("<<ComboboxSelected>>", lambda _event: self.save_config())
+
+        ttk.Label(speed_frame, text="Slider (0–100%)").grid(
+            row=1, column=0, sticky="w", padx=10, pady=5
+        )
+        speed_slider = ttk.Scale(
+            speed_frame,
+            from_=0,
+            to=100,
+            orient="horizontal",
+            variable=self.speed_slider_var,
+            command=lambda _value: self.save_config(),
+        )
+        speed_slider.grid(row=1, column=1, sticky="ew", padx=10, pady=5)
+
+        ttk.Checkbutton(
+            speed_frame,
+            text="Modo automático (ajusta se errar)",
+            variable=self.auto_speed_var,
+            command=self.save_config,
+        ).grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=5)
+
+        speed_frame.columnconfigure(1, weight=1)
+
+        control_frame = ttk.Frame(self)
+        control_frame.pack(fill="x", padx=10, pady=10)
+
+        self.start_button = ttk.Button(control_frame, text="Iniciar", command=self.start_processing)
+        self.start_button.pack(side="left", padx=5)
+
+        self.pause_button = ttk.Button(control_frame, text="Pausar", command=self.toggle_pause)
+        self.pause_button.pack(side="left", padx=5)
+
+        self.stop_button = ttk.Button(control_frame, text="Parar", command=self.stop_processing)
+        self.stop_button.pack(side="left", padx=5)
+
+    def select_file(self):
+        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
+        if not file_path:
+            return
+        self.file_path = file_path
+        self.file_label.config(text=os.path.basename(file_path))
+        self.update_columns()
+        self.update_total_lines()
+        self.status_label.config(text="Status: Aguardando")
+
+    def update_columns(self):
+        columns = []
+        if self.file_path:
+            has_header, delimiter = detect_header(self.file_path)
+            self.has_header = has_header
+            with open(self.file_path, "r", encoding="utf-8-sig", newline="") as file:
+                reader = csv.reader(file, delimiter=delimiter)
+                first_row = next(reader, [])
+                if has_header:
+                    columns = [col.strip() for col in first_row] or []
+        if not columns:
+            columns = [str(i) for i in range(1, 51)]
+            self.has_header = False
+        self.column_values = columns
+
+        for field, (var, combo) in self.column_vars.items():
+            combo["values"] = [""] + columns
+            saved = self.config.column_mapping.get(field)
+            if saved in columns:
+                var.set(saved)
+            elif self.has_header:
+                normalized_field = normalize_text(field)
+                best = next(
+                    (col for col in columns if normalize_text(col) == normalized_field),
+                    None,
+                )
+                if best is None:
+                    best = next(
+                        (col for col in columns if normalized_field in normalize_text(col)),
+                        None,
+                    )
+                if best is not None:
+                    var.set(best)
+                else:
+                    var.set("")
+            elif columns:
+                var.set("")
+
+        self.save_config()
+
+    def on_column_search(self, event, field):
+        value = event.widget.get()
+        if value == "":
+            filtered = self.column_values
+        else:
+            value_lower = value.lower()
+            filtered = [
+                column
+                for column in self.column_values
+                if value_lower in str(column).lower()
+            ]
+        _search_var, _search_entry = self.column_search_vars.get(field, (None, None))
+        _var, combo = self.column_vars.get(field, (None, None))
+        if combo is not None:
+            combo["values"] = [""] + filtered
+        self.save_config()
+
+    def update_total_lines(self):
+        if not self.file_path:
+            self.total_lines = 0
+            self.total_label.config(text="Total de linhas: 0")
+            return
+        try:
+            self.total_lines = count_lines(self.file_path, self.skip_var.get())
+            self.total_label.config(text=f"Total de linhas: {self.total_lines}")
+        except Exception as exc:
+            logging.exception("Erro ao contar linhas: %s", exc)
+            messagebox.showerror("Erro", "Não foi possível ler o arquivo CSV.")
+
+    def save_config(self):
+        mapping = {}
+        for field, (var, _combo) in self.column_vars.items():
+            if var.get():
+                mapping[field] = var.get()
+        self.config = AppConfig(
+            column_mapping=mapping,
+            skip_first_line=self.skip_var.get(),
+            speed_preset=self.speed_preset_var.get(),
+            speed_slider=int(self.speed_slider_var.get()),
+            auto_speed=self.auto_speed_var.get(),
+        )
+        save_config(self.config)
+
+    def on_skip_toggle(self):
+        self.save_config()
+        self.update_total_lines()
+
+    def resolve_mapping(self):
+        if not self.file_path:
+            return {}
+        has_header, delimiter = detect_header(self.file_path)
+        header = []
+        if has_header:
+            with open(self.file_path, "r", encoding="utf-8-sig", newline="") as file:
+                reader = csv.reader(file, delimiter=delimiter)
+                header = next(reader, [])
+        mapping = {}
+        for field, value in self.config.column_mapping.items():
+            if value in header:
+                mapping[field] = header.index(value)
+            else:
+                try:
+                    mapping[field] = int(value) - 1
+                except Exception:
+                    mapping[field] = 0
+        return mapping
+
+    def start_processing(self):
+        if not self.file_path:
+            messagebox.showwarning("Aviso", "Selecione um arquivo CSV primeiro.")
+            return
+        if self.processing_thread and self.processing_thread.is_alive():
+            messagebox.showinfo("Info", "Processamento já está em andamento.")
+            return
+        if pyautogui is None:
+            messagebox.showerror(
+                "Erro",
+                "pyautogui não está instalado. Instale para habilitar a automação.",
+            )
+            return
+        self.stop_event.clear()
+        self.pause_event.clear()
+        self.progress["value"] = 0
+        mapping = self.resolve_mapping()
+        if not mapping or any(field not in mapping for field in REQUIRED_FIELDS):
+            messagebox.showwarning("Aviso", "Configure todas as colunas antes de iniciar.")
+            return
+        self.set_controls_state("preparing")
+        self.show_preparation(mapping)
+
+    def _update_countdown(self, mapping):
+        if self.stop_event.is_set():
+            self._reset_countdown_ui(status="Status: Parado")
+            self.set_controls_state("idle")
+            return
+        if self.pause_event.is_set():
+            self.status_label.config(text="Status: Pausado (contagem)")
+            self.countdown_label.config(text=f"Início em: {self.countdown_remaining}s")
+            self.countdown_after_id = self.after(200, lambda: self._update_countdown(mapping))
+            return
+        if self.countdown_remaining <= 0:
+            self.status_label.config(text="Status: Processando")
+            self.countdown_label.config(text="Início em: 0s")
+            self.processing_thread = threading.Thread(
+                target=self._run_processing,
+                args=(mapping,),
+                daemon=True,
+            )
+            self.processing_thread.start()
+            self.set_controls_state("processing")
+            return
+        self.status_label.config(text=f"Status: Iniciando em {self.countdown_remaining}...")
+        self.countdown_label.config(text=f"Início em: {self.countdown_remaining}s")
+        self.countdown_remaining -= 1
+        self.countdown_after_id = self.after(1000, lambda: self._update_countdown(mapping))
+
+    def _run_processing(self, mapping):
+        try:
+            process_file(
+                self.file_path,
+                mapping,
+                self.skip_var.get(),
+                self.update_progress,
+                self.pause_event,
+                self.stop_event,
+                self.get_speed_settings(),
+            )
+            self.after(0, lambda: self.status_label.config(text="Status: Finalizado"))
+            self.after(0, lambda: self._reset_countdown_ui())
+            self.after(0, lambda: self.set_controls_state("idle"))
+        except Exception as exc:
+            logging.exception("Erro durante processamento: %s", exc)
+            self.after(0, lambda: messagebox.showerror("Erro", str(exc)))
+            self.after(0, lambda: self.status_label.config(text="Status: Erro"))
+            self.after(0, lambda: self._reset_countdown_ui())
+            self.after(0, lambda: self.set_controls_state("idle"))
+
+    def update_progress(self, processed, total):
+        self.after(
+            0,
+            lambda: self._update_progress_ui(processed, total),
+        )
+
+    def _update_progress_ui(self, processed, total):
+        self.progress["maximum"] = max(total, 1)
+        self.progress["value"] = processed
+        self.total_label.config(text=f"Total de linhas: {total}")
+
+    def toggle_pause(self):
+        in_countdown = self.countdown_remaining > 0 and (
+            self.processing_thread is None or not self.processing_thread.is_alive()
+        )
+        in_processing = self.processing_thread is not None and self.processing_thread.is_alive()
+
+        if not (in_countdown or in_processing):
+            return
+
+        if self.pause_event.is_set():
+            self.pause_event.clear()
+            self.pause_button.config(text="Pausar")
+            if in_countdown:
+                self.status_label.config(text=f"Status: Iniciando em {self.countdown_remaining}...")
+            else:
+                self.status_label.config(text="Status: Processando")
+        else:
+            self.pause_event.set()
+            self.pause_button.config(text="Continuar")
+            if in_countdown:
+                self.status_label.config(text="Status: Pausado (contagem)")
+            else:
+                self.status_label.config(text="Status: Pausado")
+
+    def stop_processing(self):
+        if self.prep_window is not None:
+            self.stop_event.set()
+            self.prep_window.destroy()
+            self.prep_window = None
+            self._reset_countdown_ui(status="Status: Parado")
+            self.set_controls_state("idle")
+            return
+        if self.countdown_remaining > 0:
+            self.stop_event.set()
+            if self.countdown_after_id is not None:
+                self.after_cancel(self.countdown_after_id)
+                self.countdown_after_id = None
+            self._reset_countdown_ui(status="Status: Parado")
+            self.set_controls_state("idle")
+            return
+        if not self.processing_thread or not self.processing_thread.is_alive():
+            return
+        self.stop_event.set()
+        self.pause_event.clear()
+        self.pause_button.config(text="Pausar")
+        self.status_label.config(text="Status: Parando")
+
+    def _reset_countdown_ui(self, status=None):
+        self.countdown_remaining = 0
+        self.countdown_label.config(text="Início em: -")
+        if status is not None:
+            self.status_label.config(text=status)
+
+    def get_speed_settings(self):
+        preset = self.speed_preset_var.get()
+        if preset == "Lenta":
+            base_key = 0.08
+            base_delay = 0.4
+        elif preset == "Rápida":
+            base_key = 0.01
+            base_delay = 0.1
+        else:
+            base_key = 0.04
+            base_delay = 0.25
+        slider_ratio = max(0.0, min(self.speed_slider_var.get() / 100.0, 1.0))
+        multiplier = 1.5 - slider_ratio
+        return {
+            "key_interval": base_key * multiplier,
+            "field_delay": base_delay * multiplier,
+            "auto_speed": self.auto_speed_var.get(),
+        }
+
+    def show_preparation(self, mapping):
+        if self.prep_window is not None:
+            return
+        self.prep_window = tk.Toplevel(self)
+        self.prep_window.title("Preparação")
+        self.prep_window.geometry("360x180")
+        self.prep_window.transient(self)
+        self.prep_window.grab_set()
+        message = (
+            "Preparação:\n\n"
+            "1) Clique no campo Objeto no sistema.\n"
+            "2) Volte para esta janela.\n"
+            "3) Pressione F8 para confirmar o foco.\n\n"
+            "ESC ou Parar cancelam."
+        )
+        ttk.Label(self.prep_window, text=message, justify="left").pack(
+            fill="x", padx=10, pady=10
+        )
+        ttk.Button(self.prep_window, text="Confirmar (F8)", command=lambda: self.finish_preparation(mapping)).pack(
+            pady=5
+        )
+        self.prep_window.bind("<F8>", lambda _event: self.finish_preparation(mapping))
+        self.prep_window.protocol("WM_DELETE_WINDOW", self.stop_processing)
+
+    def finish_preparation(self, mapping):
+        if self.prep_window is not None:
+            self.prep_window.destroy()
+            self.prep_window = None
+        self.set_controls_state("countdown")
+        if self.countdown_after_id is not None:
+            self.after_cancel(self.countdown_after_id)
+            self.countdown_after_id = None
+        self.countdown_remaining = self.countdown_seconds
+        self.status_label.config(text=f"Status: Iniciando em {self.countdown_remaining}...")
+        self.countdown_label.config(text=f"Início em: {self.countdown_remaining}s")
+        self.update_idletasks()
+        self._update_countdown(mapping)
+
+    def set_controls_state(self, state):
+        if state == "idle":
+            self.start_button.config(state="normal")
+            self.pause_button.config(state="disabled", text="Pausar")
+            self.stop_button.config(state="disabled")
+        elif state == "preparing":
+            self.start_button.config(state="disabled")
+            self.pause_button.config(state="disabled", text="Pausar")
+            self.stop_button.config(state="normal")
+        elif state == "countdown":
+            self.start_button.config(state="disabled")
+            self.pause_button.config(state="normal")
+            self.stop_button.config(state="normal")
+        elif state == "processing":
+            self.start_button.config(state="disabled")
+            self.pause_button.config(state="normal")
+            self.stop_button.config(state="normal")
+
+
+if __name__ == "__main__":
+    app = App()
+    app.mainloop()
 
EOF
)
