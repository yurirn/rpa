# src/ui/main_window.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from src.core.logger import set_logger_callback
import importlib
import json
import os
import threading
import base64

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

        self.selected_module_id = tk.StringVar()
        self.selected_module_name = tk.StringVar()
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.show_password = tk.BooleanVar(value=False)
        self.save_credentials = tk.BooleanVar(value=False)
        self.excel_file_path = tk.StringVar()
        self.tipo_busca = tk.StringVar(value="numero_exame")
        self.gera_xml_tiss = tk.StringVar(value="sim")
        self.headless_mode = tk.BooleanVar(value=True)
        self.pular_para_laudos = tk.BooleanVar(value=False)

        self.unimed_user = tk.StringVar()
        self.unimed_password = tk.StringVar()
        self.show_unimed_password = tk.BooleanVar(value=False)
        self.save_unimed_credentials = tk.BooleanVar(value=False)

        self.cobrar_de = tk.StringVar(value="C")
        self.data_tipo = tk.StringVar(value="recepcao")

        self.modules = self.load_modules()
        self.module_id_map = {m['id']: m for m in self.modules}
        self.module_name_map = {m['name']: m for m in self.modules}
        self.load_last_credentials()
        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.log("Sistema iniciado", "SUCCESS")
        self.set_initial_focus()
        set_logger_callback(self.log)

        self.execution_thread = None
        self.cancel_requested = threading.Event()

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

        self.headless_check = ttk.Checkbutton(
            frame,
            text="Executar em segundo plano (modo headless)",
            variable=self.headless_mode
        )
        self.headless_check.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))

        self.save_credentials_check = ttk.Checkbutton(
            frame,
            text="Salvar credenciais neste computador",
            variable=self.save_credentials,
            command=self.on_save_credentials_changed
        )
        self.save_credentials_check.grid(row=4, column=0, columnspan=2, sticky="w", pady=(5, 0))

    def create_module_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Módulo de Automação", padding="10")
        frame.grid(row=1, column=1, sticky="nsew", padx=(5, 0), pady=(0, 10))
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Selecione:").grid(row=0, column=0, sticky="w")
        self.module_combo = ttk.Combobox(
            frame,
            textvariable=self.selected_module_name,
            state="readonly",
            values=[m["name"] for m in self.modules]
        )
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

        module_id = self.selected_module_id.get()
        module = self.module_id_map.get(module_id)
        requires_excel = module.get("requires_excel") if module else False
        tipo_busca = module.get("tipo_busca") if module else False
        has_gera_xml_tiss = module.get("has_gera_xml_tiss") if module else False
        requires_unimed_credentials = module.get("requires_unimed_credentials") if module else False
        has_cobrar_de = module.get("has_cobrar_de") if module else False
        has_pular_para_laudos = module.get("has_pular_para_laudos") if module else False
        has_data_tipo = module.get("has_data_tipo") if module else False

        # Armazenar referências dos módulos para uso posterior
        self.current_module_config = {
            'requires_excel': requires_excel,
            'tipo_busca': tipo_busca,
            'has_gera_xml_tiss': has_gera_xml_tiss,
            'requires_unimed_credentials': requires_unimed_credentials,
            'has_cobrar_de': has_cobrar_de,
            'has_pular_para_laudos': has_pular_para_laudos,
            'has_data_tipo': has_data_tipo
        }

        row = 0
        if requires_excel:
            self.params_frame.columnconfigure(1, weight=1)
            ttk.Label(self.params_frame, text="Arquivo Excel:").grid(row=row, column=0, sticky="w")
            entry = ttk.Entry(self.params_frame, textvariable=self.excel_file_path)
            entry.grid(row=row, column=1, sticky="ew", padx=(0, 10))
            ttk.Button(self.params_frame, text="Selecionar", command=self.select_excel_file).grid(row=row, column=2)
            row += 1
        if tipo_busca:
            ttk.Label(self.params_frame, text="Tipo de Busca:").grid(row=row, column=0, sticky="w", pady=(15, 5))
            busca_frame = ttk.Frame(self.params_frame)
            busca_frame.grid(row=row, column=1, sticky="w", columnspan=2)
            ttk.Radiobutton(busca_frame, text="Número de Exame", variable=self.tipo_busca, value="numero_exame", command=self.on_tipo_busca_changed).pack(side=tk.LEFT, padx=(0, 20))
            ttk.Radiobutton(busca_frame, text="Número de Guia", variable=self.tipo_busca, value="numero_guia", command=self.on_tipo_busca_changed).pack(side=tk.LEFT)
            row += 1
            self.descricao_tipo = ttk.Label(self.params_frame, text=self.get_descricao_tipo_busca("numero_exame"),
                                            foreground="blue")
            self.descricao_tipo.grid(row=row, column=0, columnspan=3, sticky="w", pady=(5, 0))
            row += 1

        if has_data_tipo:
            ttk.Label(self.params_frame, text="Tipo de Data:").grid(row=row, column=0, sticky="w", pady=(15, 5))
            data_frame = ttk.Frame(self.params_frame)
            data_frame.grid(row=row, column=1, sticky="w", columnspan=2)
            ttk.Radiobutton(data_frame, text="Data Recepção", variable=self.data_tipo, value="recepcao",
                            command=self.on_data_tipo_changed).pack(side=tk.LEFT, padx=(0, 20))
            ttk.Radiobutton(data_frame, text="Data Liberação", variable=self.data_tipo, value="liberacao",
                            command=self.on_data_tipo_changed).pack(side=tk.LEFT)
            row += 1
            self.descricao_data = ttk.Label(self.params_frame, text=self.get_descricao_data_tipo("recepcao"),
                                            foreground="blue")
            self.descricao_data.grid(row=row, column=0, columnspan=3, sticky="w", pady=(5, 0))
            row += 1

        if has_cobrar_de:
            ttk.Label(self.params_frame, text="Cobrar de:").grid(row=row, column=0, sticky="w", pady=(15, 5))
            cobrar_frame = ttk.Frame(self.params_frame)
            cobrar_frame.grid(row=row, column=1, sticky="w", columnspan=2)
            ttk.Radiobutton(cobrar_frame, text="Convênio", variable=self.cobrar_de, value="C",
                            command=self.on_cobrar_de_changed).pack(side=tk.LEFT, padx=(0, 20))
            ttk.Radiobutton(cobrar_frame, text="Procedência", variable=self.cobrar_de, value="P",
                            command=self.on_cobrar_de_changed).pack(side=tk.LEFT)
            row += 1
            self.descricao_cobrar = ttk.Label(self.params_frame, text=self.get_descricao_cobrar_de("C"),
                                              foreground="blue")
            self.descricao_cobrar.grid(row=row, column=0, columnspan=3, sticky="w", pady=(5, 0))
            row += 1

        if has_gera_xml_tiss:
            ttk.Label(self.params_frame, text="Gera XML TISS?:").grid(row=row, column=0, sticky="w", pady=(15, 5))
            gera_xml_frame = ttk.Frame(self.params_frame)
            gera_xml_frame.grid(row=row, column=1, sticky="w", columnspan=2)
            ttk.Radiobutton(gera_xml_frame, text="Sim", variable=self.gera_xml_tiss, value="sim", command=self.on_gera_xml_tiss_changed).pack(side=tk.LEFT, padx=(0, 20))
            ttk.Radiobutton(gera_xml_frame, text="Não", variable=self.gera_xml_tiss, value="nao", command=self.on_gera_xml_tiss_changed).pack(side=tk.LEFT)
            row += 1

        if has_pular_para_laudos:
            self.pular_para_laudos_check = ttk.Checkbutton(
                self.params_frame,
                text="⚡ Pular processo de conclusão e ir direto para visualização de laudos",
                variable=self.pular_para_laudos
            )
            self.pular_para_laudos_check.grid(row=row, column=0, columnspan=3, sticky="w", pady=(15, 0))
            row += 1

        # Criar container para credenciais Unimed (será mostrado/escondido dinamicamente)
        if requires_unimed_credentials:
            self.unimed_credentials_frame = ttk.Frame(self.params_frame)
            self.unimed_credentials_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(10, 0))
            self.unimed_credentials_frame.columnconfigure(1, weight=1)

            ttk.Label(self.unimed_credentials_frame, text="Credenciais Unimed:", font=('Arial', 10, 'bold')).grid(row=0, column=0, columnspan=3, sticky="w")

            ttk.Label(self.unimed_credentials_frame, text="Usuário Unimed:").grid(row=1, column=0, sticky="w", pady=(5, 0))
            unimed_user_entry = ttk.Entry(self.unimed_credentials_frame, textvariable=self.unimed_user, width=30)
            unimed_user_entry.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=(5, 0))

            ttk.Label(self.unimed_credentials_frame, text="Senha Unimed:").grid(row=2, column=0, sticky="w", pady=(5, 0))
            unimed_pass_frame = ttk.Frame(self.unimed_credentials_frame)
            unimed_pass_frame.grid(row=2, column=1, sticky="w", padx=(10, 0), pady=(5, 0))

            self.unimed_password_entry = ttk.Entry(unimed_pass_frame, textvariable=self.unimed_password, show="*", width=25)
            self.unimed_password_entry.pack(side=tk.LEFT)

            ttk.Checkbutton(unimed_pass_frame, text="Mostrar", variable=self.show_unimed_password, command=self.toggle_unimed_password_visibility).pack(side=tk.LEFT, padx=(10, 0))

            # Checkbox para salvar credenciais da Unimed
            self.save_unimed_check = ttk.Checkbutton(
                self.unimed_credentials_frame,
                text="Salvar credenciais Unimed neste computador",
                variable=self.save_unimed_credentials,
                command=self.on_save_unimed_credentials_changed
            )
            self.save_unimed_check.grid(row=3, column=0, columnspan=3, sticky="w", pady=(5, 0))

            # Inicialmente mostrar/esconder baseado no valor atual do gera_xml_tiss
            self.update_unimed_credentials_visibility()
            row += 1

        if not requires_excel and not has_gera_xml_tiss:
            ttk.Label(self.params_frame, text="Os parâmetros aparecerão aqui quando um módulo for selecionado", foreground="gray").pack(pady=20)

    def on_tipo_busca_changed(self):
        self.descricao_tipo.config(text=self.get_descricao_tipo_busca(self.tipo_busca.get()))

    def on_gera_xml_tiss_changed(self):
        """Callback chamado quando a opção 'Gera XML TISS' é alterada"""
        self.update_unimed_credentials_visibility()

    def on_data_tipo_changed(self):
        """Callback chamado quando a opção 'Tipo de Data' é alterada"""
        self.descricao_data.config(text=self.get_descricao_data_tipo(self.data_tipo.get()))

    def on_cobrar_de_changed(self):
        """Callback chamado quando a opção 'Cobrar de' é alterada"""
        self.descricao_cobrar.config(text=self.get_descricao_cobrar_de(self.cobrar_de.get()))

    def get_descricao_data_tipo(self, tipo):
        """Retorna a descrição do tipo de data selecionado"""
        if tipo == "recepcao":
            return "📅 Busca será feita pela data de recepção do exame."
        return "📅 Busca será feita pela data de liberação do exame."

    def get_descricao_cobrar_de(self, tipo):
        """Retorna a descrição do tipo de cobrança selecionado"""
        if tipo == "C":
            return "💳 Cobrança será direcionada ao convênio."
        return "🏥 Cobrança será direcionada à procedência."

    def update_unimed_credentials_visibility(self):
        """Mostra ou esconde as credenciais da Unimed baseado na seleção do XML TISS"""
        if hasattr(self, 'unimed_credentials_frame') and hasattr(self, 'current_module_config'):
            if (self.current_module_config.get('requires_unimed_credentials') and
                self.current_module_config.get('has_gera_xml_tiss')):

                if self.gera_xml_tiss.get() == "sim":
                    # Mostrar credenciais da Unimed
                    self.unimed_credentials_frame.grid()
                else:
                    # Esconder credenciais da Unimed
                    self.unimed_credentials_frame.grid_remove()

    def get_descricao_tipo_busca(self, tipo):
        if tipo == "numero_exame":
            return "📋 Busca pelo numero do exame."
        return "📋 Busca pelo numero da guia."

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[["Excel files", "*.xlsx"]])
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

        self.log_text = tk.Text(frame, wrap="word", height=20, font=("Consolas", 9))
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
        file_menu.add_command(label="Limpar Credenciais Salvas", command=self.clear_saved_credentials)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.root.quit)

    def toggle_password_visibility(self):
        self.password_entry.config(show="" if self.show_password.get() else "*")

    def toggle_unimed_password_visibility(self):
        self.unimed_password_entry.config(show="" if self.show_unimed_password.get() else "*")

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
        # Recarrega módulos do arquivo para refletir alterações recentes (ex.: requires_excel)
        self.modules = self.load_modules()
        self.module_id_map = {m['id']: m for m in self.modules}
        self.module_name_map = {m['name']: m for m in self.modules}

        # Buscar pelo nome selecionado e obter o id
        selected_name = self.module_combo.get()
        module = self.module_name_map.get(selected_name)
        if not module:
            self.module_description.config(text="Módulo não encontrado")
            return
        self.selected_module_id.set(module['id'])
        self.selected_module_name.set(module['name'])
        desc = module.get("description", "Sem descrição disponível")
        self.module_description.config(text=f"📋 {desc}")
        self.run_button.config(state="normal")
        self.update_params_section()
        self.log(f"Módulo selecionado: {selected_name} (id: {module['id']})", "INFO")

    def run_module(self):
        module_id = self.selected_module_id.get()
        if not module_id or not self.username.get().strip() or not self.password.get().strip():
            messagebox.showwarning("Aviso", "Preencha usuário, senha e selecione o módulo.")
            return
        module = self.module_id_map.get(module_id)
        if not module:
            self.log("Módulo não encontrado", "ERROR")
            return
        params = {
            "username": self.username.get(),
            "password": self.password.get(),
            "cancel_flag": self.cancel_requested,
            "headless_mode": self.headless_mode.get()
        }
        if module.get("requires_excel"):
            excel_path = self.excel_file_path.get()
            if not os.path.exists(excel_path):
                messagebox.showerror("Erro", "Arquivo Excel não encontrado!")
                return
            params.update({
                "excel_file": excel_path,
                "modo_busca": "exame" if self.tipo_busca.get() == "numero_exame" else "guia"
            })
        if module.get("has_gera_xml_tiss"):
            params["gera_xml_tiss"] = self.gera_xml_tiss.get()
        if module.get("has_cobrar_de"):
            params["cobrar_de"] = self.cobrar_de.get()
        if module.get("has_pular_para_laudos"):
            params["pular_para_laudos"] = self.pular_para_laudos.get()
        if module.get("has_data_tipo"):
            params["data_tipo"] = self.data_tipo.get()
        if module.get("requires_unimed_credentials"):
            unimed_user = self.unimed_user.get().strip()
            unimed_pass = self.unimed_password.get().strip()
            if not unimed_user or not unimed_pass:
                if self.gera_xml_tiss.get() == "sim" or self.gera_xml_tiss.get() is None:
                    messagebox.showwarning("Credenciais Unimed", "Preencha usuário e senha da Unimed!")
                    return
            params.update({
                "unimed_user": unimed_user,
                "unimed_pass": unimed_pass
            })
        def run_in_thread():
            try:
                mod = importlib.import_module(module["module_path"]) 
                if hasattr(mod, "run"):
                    mod.run(params)
            except Exception as e:
                self.log(f"Erro: {e}", "ERROR")
            finally:
                self.root.after(0, self.execution_finished)
        self.run_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.cancel_requested.clear()
        self.execution_thread = threading.Thread(target=run_in_thread, daemon=True)
        self.execution_thread.start()

    def execution_finished(self):
        self.log("Execução finalizada", "SUCCESS")
        self.run_button.config(state="normal")
        self.stop_button.config(state="disabled")

    def stop_module(self):
        self.log("Execução interrompida pelo usuário", "WARNING")
        self.cancel_requested.set()
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

    def _encode_password(self, password):
        """Codifica a senha usando base64 para armazenamento local"""
        if not password:
            return ""
        # Adiciona um salt simples baseado no username para ofuscar melhor
        username = self.username.get().strip()
        salt = f"{username}_rpa_salt"
        password_with_salt = f"{password}_{salt}"
        return base64.b64encode(password_with_salt.encode('utf-8')).decode('utf-8')

    def _decode_password(self, encoded_password):
        """Decodifica a senha armazenada"""
        if not encoded_password:
            return ""
        try:
            decoded = base64.b64decode(encoded_password.encode('utf-8')).decode('utf-8')
            # Remove o salt
            username = self.username.get().strip()
            salt = f"{username}_rpa_salt"
            if decoded.endswith(f"_{salt}"):
                return decoded[:-len(f"_{salt}")]
            return decoded
        except:
            return ""

    def _encode_unimed_password(self, password):
        """Codifica a senha da Unimed usando base64 para armazenamento local"""
        if not password:
            return ""
        # Adiciona um salt simples baseado no username da Unimed para ofuscar melhor
        unimed_user = self.unimed_user.get().strip()
        salt = f"{unimed_user}_unimed_salt"
        password_with_salt = f"{password}_{salt}"
        return base64.b64encode(password_with_salt.encode('utf-8')).decode('utf-8')

    def _decode_unimed_password(self, encoded_password):
        """Decodifica a senha da Unimed armazenada"""
        if not encoded_password:
            return ""
        try:
            decoded = base64.b64decode(encoded_password.encode('utf-8')).decode('utf-8')
            # Remove o salt
            unimed_user = self.unimed_user.get().strip()
            salt = f"{unimed_user}_unimed_salt"
            if decoded.endswith(f"_{salt}"):
                return decoded[:-len(f"_{salt}")]
            return decoded
        except:
            return ""

    def clear_logs(self):
        self.log_text.delete(1.0, tk.END)

    def clear_credentials(self):
        self.username.set("")
        self.password.set("")
        self.save_credentials.set(False)
        self.set_initial_focus()

    def clear_saved_credentials(self):
        """Remove as credenciais salvas do arquivo de configuração"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                
                # Limpar credenciais do Pathoweb
                config.pop('last_username', None)
                config.pop('last_password', None)
                config['save_credentials'] = False
                
                # Limpar credenciais da Unimed
                config.pop('last_unimed_username', None)
                config.pop('last_unimed_password', None)
                config['save_unimed_credentials'] = False
                
                with open(CONFIG_FILE, 'w') as f:
                    json.dump(config, f, indent=2)
                self.log("Todas as credenciais salvas foram removidas", "INFO")
        except Exception as e:
            self.log(f"Erro ao limpar credenciais salvas: {e}", "ERROR")

    def show_about(self):
        about_text = f"""Sistema RPA - Clínica

Versão: {APP_VERSION}
Desenvolvido para automatizar processos da clínica

Este é o layout base do sistema.
Os módulos serão implementados gradualmente.
        """
        messagebox.showinfo("Sobre", about_text)

    def load_last_credentials(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    
                    # Carregar credenciais do Pathoweb
                    last_user = config.get('last_username', '')
                    last_password_encoded = config.get('last_password', '')
                    save_credentials = config.get('save_credentials', False)

                    if last_user:
                        self.username.set(last_user)
                    if last_password_encoded and save_credentials:
                        decoded_password = self._decode_password(last_password_encoded)
                        self.password.set(decoded_password)
                    self.save_credentials.set(save_credentials)
                    
                    # Carregar credenciais da Unimed
                    last_unimed_user = config.get('last_unimed_username', '')
                    last_unimed_password_encoded = config.get('last_unimed_password', '')
                    save_unimed_credentials = config.get('save_unimed_credentials', False)
                    
                    if last_unimed_user:
                        self.unimed_user.set(last_unimed_user)
                    if last_unimed_password_encoded and save_unimed_credentials:
                        decoded_unimed_password = self._decode_unimed_password(last_unimed_password_encoded)
                        self.unimed_password.set(decoded_unimed_password)
                    self.save_unimed_credentials.set(save_unimed_credentials)
        except Exception as e:
            self.log(f"Erro ao carregar credenciais salvas: {e}", "ERROR")

    def save_last_username(self):
        try:
            config = {}
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
            
            # Sempre salva o usuário do Pathoweb
            config['last_username'] = self.username.get().strip()
            
            # Salva senha e opção apenas se o usuário escolheu salvar credenciais
            config['save_credentials'] = self.save_credentials.get()
            if self.save_credentials.get():
                config['last_password'] = self._encode_password(self.password.get().strip())
            else:
                # Remove senha salva se usuário desabilitou a opção
                config.pop('last_password', None)
            
            # Salvar credenciais da Unimed
            unimed_user = self.unimed_user.get().strip()
            if unimed_user:
                config['last_unimed_username'] = unimed_user
            
            config['save_unimed_credentials'] = self.save_unimed_credentials.get()
            if self.save_unimed_credentials.get():
                unimed_pass = self.unimed_password.get().strip()
                if unimed_pass:
                    config['last_unimed_password'] = self._encode_unimed_password(unimed_pass)
            else:
                # Remove senha da Unimed salva se usuário desabilitou a opção
                config.pop('last_unimed_password', None)
            
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            self.log(f"Erro ao salvar credenciais: {e}", "ERROR")

    def on_save_credentials_changed(self):
        if self.save_credentials.get():
            self.log("Credenciais Pathoweb serão salvas localmente", "INFO")
        else:
            self.log("Credenciais Pathoweb não serão salvas localmente", "INFO")

    def on_save_unimed_credentials_changed(self):
        if self.save_unimed_credentials.get():
            self.log("Credenciais Unimed serão salvas localmente", "INFO")
        else:
            self.log("Credenciais Unimed não serão salvas localmente", "INFO")

    def on_closing(self):
        if self.username.get().strip():
            self.save_last_username()
        self.root.destroy()

    def run(self):
        self.root.mainloop()

