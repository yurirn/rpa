import threading
from gui_interface import GUIInterface
from web_automation import WebAutomation
from login_manager import LoginManager

class AutomationSystem:
    def __init__(self):
        self.gui = GUIInterface()
        self.web_automation = WebAutomation(self.gui)
        self.login_manager = LoginManager()
        self.monitoring = False
        
        # Iniciar automaticamente
        self.auto_start()
        
    def auto_start(self):
        """Inicia automaticamente o sistema"""
        thread = threading.Thread(target=self.run_automation)
        thread.daemon = True
        thread.start()
        
    def run_automation(self):
        """Executa todo o processo de automação"""
        try:
            # Verificar configurações
            is_valid, message = self.login_manager.validate_configuration()
            if not is_valid:
                self.gui.log_message(f"✗ {message}", "error")
                return
                
            credentials = self.login_manager.get_credentials()
            
            self.gui.log_message("=== INICIANDO SISTEMA ===", "info")
            self.gui.log_message(f"URL: {credentials['system_url']}", "info")
            
            # Determinar tipo de login
            if credentials['auto_login']:
                self.gui.log_message("→ Login automático configurado", "info")
            else:
                self.gui.log_message("→ Login manual (usuário fará login)", "warning")
            
            # Configurar Chrome
            self.web_automation.setup_chrome()
            
            # Acessar sistema
            self.web_automation.access_system(credentials['system_url'])
            
            # Processo de login
            login_success = self.handle_login_process(credentials)
            
            if login_success:
                # Após login, executar ações
                self.execute_post_login_actions()
            else:
                self.gui.log_message("✗ Falha no processo de login", "error")
                
        except Exception as e:
            self.gui.log_message(f"✗ Erro: {str(e)}", "error")
            self.gui.update_status("Erro no sistema", "red")
            
    def handle_login_process(self, credentials):
        """Gerencia o processo de login (automático ou manual)"""
        if credentials['auto_login']:
            success = self.web_automation.perform_auto_login(
                credentials['username'], 
                credentials['password']
            )
            
            if not success:
                return self.web_automation.wait_manual_login()
            return success
        else:
            return self.web_automation.wait_manual_login()
            
    def execute_post_login_actions(self):
        """Executa ações após login"""
        self.gui.update_status("Executando automação...", "blue")
        self.gui.log_message("=== INICIANDO AUTOMAÇÃO ===", "info")
        
        self.web_automation.create_new_exam()
        
        self.gui.update_status("Automação concluída", "green")
        self.gui.log_message("=== AUTOMAÇÃO FINALIZADA ===", "success")
        
    def run(self):
        """Inicia a aplicação"""
        self.gui.run(self.on_closing)
        
    def on_closing(self):
        """Callback para fechamento da aplicação"""
        self.monitoring = False
        self.web_automation.quit_driver()
        self.gui.root.destroy()

if __name__ == "__main__":
    app = AutomationSystem()
    app.run()