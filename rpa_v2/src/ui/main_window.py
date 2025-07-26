# src/ui/main_window.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from src.core.logger import set_logger_callback
import importlib
import json
import os

CONFIG_FILE = 'config.json'
MODULES_FILE = 'modules.json'
APP_TITLE = "Sistema RPA - Clínica"
APP_VERSION = "1.0.0"

class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)
        self.center_window()

        self.selected_module = tk.StringVar()
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.show_password = tk.BooleanVar(value=False)
        self.excel_file_path = tk.StringVar()
        self.tipo_busca = tk.StringVar(value="numero_exame")

        self.modules = self.load_modules()
        self.load_last_username()
        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.log("Sistema iniciado", "SUCCESS")
        self.set_initial_focus()
        set_logger_callback(self.log)

    def load_modules(self):
        try:
            with open(MODULES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar módulos: {e}")
            return []

    def center_window(self):
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        ww, wh = self.root.winfo_width(), self.root.winfo_height()
        x, y = (sw - ww) // 2, (sh - wh) // 2
        self.root.geometry(f"{ww}x{wh}+{x}+{y}")

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        ttk.Label(main_frame, text="Sistema de Automacao RPA", font=('Arial', 16, 'bold')).grid(row=0, column=0, columnspan=2, pady=(0, 20))
        self.create_credentials_section(main_frame)
        self.create_module_section(main_frame)
        self.create_params_section(main_frame)
        self.create_control_buttons(main_frame)
        self.create_log_section(main_frame)
        self.create_menu()

    def create_credentials_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Credenciais do Sistema", padding="10")
        frame.grid(row=1, column=0, sticky="nsew", padx=(0, 5), pady=(0, 10))
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Usuário:").grid(row=0, column=0, sticky="w")
        self.username_entry = ttk.Entry(frame, textvariable=self.username, width=40)
        self.username_entry.grid(row=0, column=1, sticky="ew", pady=5)

        ttk.Label(frame, text="Senha:").grid(row=1, column=0, sticky="w")
        password_frame = ttk.Frame(frame)
        password_frame.grid(row=1, column=1, sticky="ew", pady=5)
        password_frame.columnconfigure(0, weight=1)

        self.password_entry = ttk.Entry(password_frame, textvariable=self.password, show="*")
        self.password_entry.grid(row=0, column=0, sticky="ew")

        self.show_password_check = ttk.Checkbutton(password_frame, text="Mostrar", variable=self.show_password, command=self.toggle_password_visibility)
        self.show_password_check.grid(row=0, column=1, padx=(10, 0))

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        self.test_login_button = ttk.Button(button_frame, text="Testar Conexão", command=self.test_connection)
        self.test_login_button.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Limpar", command=self.clear_credentials).pack(side=tk.LEFT)

    def create_module_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Módulo de Automação", padding="10")
        frame.grid(row=1, column=1, sticky="nsew", padx=(5, 0), pady=(0, 10))
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Selecione:").grid(row=0, column=0, sticky="w")
        self.module_combo = ttk.Combobox(frame, textvariable=self.selected_module, state="readonly", values=[m["name"] for m in self.modules])
        self.module_combo.grid(row=0, column=1, sticky="ew")
        self.module_combo.bind('<<ComboboxSelected>>', self.on_module_selected)

        self.module_description = ttk.Label(frame, text="Selecione um módulo para ver sua descrição", foreground="gray", wraplength=300)
        self.module_description.grid(row=1, column=0, columnspan=2, sticky="w", pady=(10, 0))

    def create_params_section(self, parent):
        self.params_frame = ttk.LabelFrame(parent, text="Parâmetros", padding="10")
        self.params_frame.grid(row=2, column=0, columnspan=2, sticky="nsew")
        self.update_params_section()

    def update_params_section(self):
        for widget in self.params_frame.winfo_children():
            widget.destroy()

        if self.selected_module.get() == "Preparação Lote (Excel)":
            self.params_frame.columnconfigure(1, weight=1)

            ttk.Label(self.params_frame, text="Arquivo Excel:").grid(row=0, column=0, sticky="w")
            entry = ttk.Entry(self.params_frame, textvariable=self.excel_file_path)
            entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))
            ttk.Button(self.params_frame, text="Selecionar", command=self.select_excel_file).grid(row=0, column=2)

            ttk.Label(self.params_frame, text="Tipo de Busca:").grid(row=1, column=0, sticky="w", pady=(15, 5))
            busca_frame = ttk.Frame(self.params_frame)
            busca_frame.grid(row=1, column=1, sticky="w", columnspan=2)
            ttk.Radiobutton(busca_frame, text="Número de Exame", variable=self.tipo_busca, value="numero_exame", command=self.on_tipo_busca_changed).pack(side=tk.LEFT, padx=(0, 20))
            ttk.Radiobutton(busca_frame, text="Número de Guia", variable=self.tipo_busca, value="numero_guia", command=self.on_tipo_busca_changed).pack(side=tk.LEFT)

            self.descricao_tipo = ttk.Label(self.params_frame, text=self.get_descricao_tipo_busca("numero_exame"), foreground="blue")
            self.descricao_tipo.grid(row=2, column=0, columnspan=3, sticky="w", pady=(5, 0))
        else:
            ttk.Label(self.params_frame, text="Os parâmetros aparecerão aqui quando um módulo for selecionado", foreground="gray").pack(pady=20)

    def on_tipo_busca_changed(self):
        self.descricao_tipo.config(text=self.get_descricao_tipo_busca(self.tipo_busca.get()))

    def get_descricao_tipo_busca(self, tipo):
        if tipo == "numero_exame":
            return "📋 Busca pelo numero do exame."
        return "📋 Busca pelo numero da guia."

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.excel_file_path.set(file_path)

    def create_control_buttons(self, parent):
        frame = ttk.Frame(parent)
        frame.grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 10))
        self.run_button = ttk.Button(frame, text="▶ Executar", command=self.run_module)
        self.run_button.pack(side=tk.LEFT, padx=(0, 10))
        self.stop_button = ttk.Button(frame, text="■ Parar", command=self.stop_module)
        self.stop_button.pack(side=tk.LEFT)

    def create_log_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Logs", padding="10")
        frame.grid(row=4, column=0, columnspan=2, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(frame, wrap="word", height=12, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        for tag, color in [("INFO", "blue"), ("SUCCESS", "green"), ("WARNING", "orange"), ("ERROR", "red")]:
            self.log_text.tag_configure(tag, foreground=color)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Arquivo", menu=file_menu)
        file_menu.add_command(label="Limpar Logs", command=self.clear_logs)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.root.quit)

    def toggle_password_visibility(self):
        self.password_entry.config(show="" if self.show_password.get() else "*")

    def test_connection(self):
        username = self.username.get().strip()
        password = self.password.get().strip()
        if not username or not password:
            messagebox.showwarning("Aviso", "Por favor, preencha usuário e senha!")
            self.log("Credenciais incompletas", "WARNING")
            return

        self.log(f"Testando conexão com usuário: {username}", "INFO")
        self.test_login_button.config(state="disabled", text="Testando...")
        self.root.after(1000, self._finish_connection_test)

    def _finish_connection_test(self):
        self.test_login_button.config(state="normal", text="Testar Conexão")
        self.log("✅ Credenciais salvas (teste será implementado com Selenium)", "SUCCESS")
        messagebox.showinfo("Sucesso", "Credenciais salvas com sucesso!")

    def validate_credentials(self):
        username = self.username.get().strip()
        password = self.password.get().strip()
        if not username or not password:
            messagebox.showwarning(
                "Credenciais Necessárias",
                "Por favor, preencha usuário e senha antes de executar!"
            )
            return False
        return True

    def on_module_selected(self, event=None):
        module_name = self.selected_module.get()
        if not module_name:
            return
        desc = next((m["description"] for m in self.modules if m["name"] == module_name), "Sem descrição disponível")
        self.module_description.config(text=f"📋 {desc}")
        self.run_button.config(state="normal")
        self.update_params_section()
        self.log(f"Módulo selecionado: {module_name}", "INFO")

    def run_module(self):
        module_name = self.selected_module.get()
        if not module_name or not self.username.get().strip() or not self.password.get().strip():
            messagebox.showwarning("Aviso", "Preencha usuário, senha e selecione o módulo.")
            return

        module_info = next((m for m in self.modules if m["name"] == module_name), None)
        if not module_info:
            self.log("Módulo não encontrado", "ERROR")
            return

        params = {
            "username": self.username.get(),
            "password": self.password.get()
        }

        if module_name == "Preparação Lote (Excel)":
            excel_path = self.excel_file_path.get()
            if not os.path.exists(excel_path):
                messagebox.showerror("Erro", "Arquivo Excel não encontrado!")
                return
            params.update({
                "excel_file": excel_path,
                "modo_busca": "exame" if self.tipo_busca.get() == "numero_exame" else "guia"
            })

        try:
            mod = importlib.import_module(module_info["module_path"])
            if hasattr(mod, "run"):
                mod.run(params)
        except Exception as e:
            self.log(f"Erro: {e}", "ERROR")

        self.run_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.root.after(2000, self.execution_finished)

    def execution_finished(self):
        self.log("Execução finalizada (simulação)", "SUCCESS")
        self.run_button.config(state="normal")
        self.stop_button.config(state="disabled")

    def stop_module(self):
        self.log("Execução interrompida pelo usuário", "WARNING")
        self.run_button.config(state="normal")
        self.stop_button.config(state="disabled")

    def set_initial_focus(self):
        (self.password_entry if self.username.get() else self.username_entry).focus_set()

    def log(self, message: str, level: str = "INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", level)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def get_credentials(self):
        return {
            'username': self.username.get().strip(),
            'password': self.password.get().strip()
        }

    def clear_logs(self):
        self.log_text.delete(1.0, tk.END)

    def clear_credentials(self):
        self.username.set("")
        self.password.set("")
        self.set_initial_focus()

    def show_about(self):
        about_text = f"""Sistema RPA - Clínica

Versão: {APP_VERSION}
Desenvolvido para automatizar processos da clínica

Este é o layout base do sistema.
Os módulos serão implementados gradualmente.
        """
        messagebox.showinfo("Sobre", about_text)

    def load_last_username(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    last_user = config.get('last_username', '')
                    if last_user:
                        self.username.set(last_user)
        except Exception as e:
            self.log(f"Erro ao carregar usuário salvo: {e}", "ERROR")

    def save_last_username(self):
        try:
            config = {}
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
            config['last_username'] = self.username.get().strip()
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            self.log(f"Erro ao salvar usuário: {e}", "ERROR")

    def on_closing(self):
        if self.username.get().strip():
            self.save_last_username()
        self.root.destroy()

    def run(self):
        self.root.mainloop()