import tkinter as tk
from tkinter import ttk
import time
from config import GUI_TITLE, GUI_GEOMETRY, GUI_FONT_TITLE, GUI_FONT_STATUS, GUI_FONT_LOG, LOG_COLORS

class GUIInterface:
    def __init__(self):
        self.root = None
        self.status_label = None
        self.log_text = None
        self.setup_gui()
        
    def setup_gui(self):
        """Configura a interface gráfica"""
        self.root = tk.Tk()
        self.root.title(GUI_TITLE)
        self.root.geometry(GUI_GEOMETRY)
        
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # Título
        title_label = ttk.Label(main_frame, text="Sistema de Automação - Monitor de Logs", 
                               font=GUI_FONT_TITLE)
        title_label.pack(pady=(0, 10))
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Iniciando sistema...", 
                                    foreground="orange", font=GUI_FONT_STATUS)
        self.status_label.pack(pady=(0, 10))
        
        # Log
        log_frame = ttk.LabelFrame(main_frame, text="Log de Atividades", padding="10")
        log_frame.pack(fill="both", expand=True)
        
        text_frame = ttk.Frame(log_frame)
        text_frame.pack(fill="both", expand=True)
        
        self.log_text = tk.Text(text_frame, wrap="word", font=GUI_FONT_LOG)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Configurar cores
        for log_type, color in LOG_COLORS.items():
            self.log_text.tag_configure(log_type, foreground=color)
        
    def log_message(self, message, type="info"):
        """Adiciona mensagem ao log com timestamp"""
        timestamp = time.strftime('%H:%M:%S')
        full_message = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, full_message, type)
        self.log_text.see(tk.END)
        self.root.update()
        
    def update_status(self, message, color="blue"):
        """Atualiza o status na interface"""
        self.status_label.config(text=message, foreground=color)
        
    def run(self, on_closing_callback):
        """Inicia a interface gráfica"""
        self.root.protocol("WM_DELETE_WINDOW", on_closing_callback)
        self.root.mainloop()