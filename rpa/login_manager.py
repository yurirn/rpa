import os
from dotenv import load_dotenv

class LoginManager:
    def __init__(self):
        load_dotenv()
        
    def get_credentials(self):
        """Obtém credenciais e configurações do .env"""
        system_url = os.getenv('SYSTEM_URL')
        login_user = os.getenv('LOGIN_USER', '')
        login_pass = os.getenv('LOGIN_PASS', '')
        
        return {
            'system_url': system_url,
            'username': login_user,
            'password': login_pass,
            'auto_login': bool(login_user and login_pass)
        }
        
    def validate_configuration(self):
        """Valida se as configurações necessárias estão presentes"""
        credentials = self.get_credentials()
        
        if not credentials['system_url']:
            return False, "SYSTEM_URL não configurada no .env"
            
        return True, "Configuração válida"