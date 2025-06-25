from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from config import CHROME_OPTIONS, SELECTORS, LOGIN_SUCCESS_PATTERNS, LOGIN_FAIL_PATTERNS, TIMEOUTS, PATIENT_NAME

class WebAutomation:
    def __init__(self, gui_interface):
        self.driver = None
        self.gui = gui_interface
        self.monitoring = False
        
    def setup_chrome(self):
        """Configura e inicia Chrome"""
        self.gui.log_message("Configurando Chrome...", "info")
        chrome_options = Options()
        
        for option in CHROME_OPTIONS:
            chrome_options.add_argument(option)
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.gui.log_message("✓ Chrome iniciado", "success")
        
    def access_system(self, url):
        """Acessa o sistema"""
        self.gui.log_message(f"Acessando: {url}", "info")
        self.driver.get(url)
        
    def perform_auto_login(self, username, password):
        """Executa login automático"""
        try:
            self.gui.update_status("Fazendo login automático...", "orange")
            self.gui.log_message("=== LOGIN AUTOMÁTICO ===", "info")
            
            # Aguardar campos de login
            user_field = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.presence_of_element_located((By.ID, SELECTORS['username_field']))
            )
            pass_field = self.driver.find_element(By.ID, SELECTORS['password_field'])
            
            # Preencher campos
            user_field.clear()
            user_field.send_keys(username)
            time.sleep(1)
            pass_field.clear()
            pass_field.send_keys(password)
            
            # Buscar botão de login
            login_button = self.driver.find_element(By.XPATH, SELECTORS['login_button'])
            login_button.click()
            
            self.gui.log_message("✓ Credenciais enviadas", "success")
            
            # Aguardar redirecionamento
            time.sleep(TIMEOUTS['login_redirect'])
            if self.check_login_success():
                self.gui.log_message("✓ Login automático realizado!", "success")
                self.gui.update_status("Login realizado", "green")
                return True
            else:
                self.gui.log_message("✗ Falha no login automático", "error")
                return False
                
        except Exception as e:
            self.gui.log_message(f"✗ Erro no login automático: {str(e)}", "error")
            return False
            
    def wait_manual_login(self):
        """Aguarda login manual do usuário"""
        self.gui.update_status("Aguardando login manual...", "orange")
        self.gui.log_message("=== AGUARDANDO LOGIN MANUAL ===", "warning")
        self.gui.log_message("→ Faça login manualmente no Chrome", "warning")
        
        initial_url = self.driver.current_url
        self.monitoring = True
        
        while self.monitoring:
            try:
                current_url = self.driver.current_url
                if current_url != initial_url and self.check_login_success():
                    self.gui.log_message("✓ Login manual detectado!", "success")
                    self.gui.update_status("Login realizado", "green")
                    return True
                time.sleep(TIMEOUTS['manual_login_check'])
            except:
                break
        return False
                
    def check_login_success(self):
        """Verifica se login foi bem-sucedido"""
        try:
            current_url = self.driver.current_url.lower()        
            
            success_indicators = [pattern in current_url for pattern in LOGIN_SUCCESS_PATTERNS]
            login_indicators = [pattern in current_url for pattern in LOGIN_FAIL_PATTERNS]
            
            return any(success_indicators) and not any(login_indicators)
        except:
            return False
            
    def create_new_exam(self):
        """Processo completo de criação de exame"""
        try:
            self.gui.log_message("→ Iniciando criação de exame", "info")
            
            # Primeiro botão
            first_button = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.element_to_be_clickable((By.XPATH, SELECTORS['create_exam_button']))
            )
            first_button.click()
            self.gui.log_message("✓ Primeiro botão clicado", "success")
            
            # Segundo botão (modal)
            time.sleep(TIMEOUTS['page_load'])
            modal_button = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.presence_of_element_located((By.XPATH, SELECTORS['modal_create_button']))
            )
            self.driver.execute_script("arguments[0].click();", modal_button)
            self.gui.log_message("✓ Botão do modal clicado", "success")
            
            self.search_patient()

            self.create_patient()

            
        except Exception as e:
            self.gui.log_message(f"✗ Erro ao criar exame: {str(e)}", "error")

    def search_patient(self):
        """Busca paciente pelo nome"""
        try:
            self.gui.log_message("→ Buscando paciente...", "info")
            time.sleep(TIMEOUTS['page_load']) 
            
            search_field = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.presence_of_element_located((By.ID, SELECTORS['patient_search_field']))
            )
            
            search_field.clear()
            search_field.send_keys(PATIENT_NAME)
            self.gui.log_message(f"✓ Nome '{PATIENT_NAME}' inserido no campo", "success")
            
            consult_button = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.presence_of_element_located((By.ID, SELECTORS['consult_button']))
            )
            self.driver.execute_script("arguments[0].click();", consult_button)
            self.gui.log_message("✓ Botão consultar clicado via JavaScript", "success")
            
            time.sleep(TIMEOUTS['search_result'])  # Aguardar resultado da busca
            self.gui.log_message("✓ Busca de paciente realizada", "success")

            create_patient = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.presence_of_element_located((By.ID, SELECTORS['create_patient']))
            )

            time.sleep(TIMEOUTS['search_result']) 
            self.driver.execute_script("arguments[0].click();", create_patient)
            self.gui.log_message("✓ Botão criar paciente clicado via Javascript", "success")

            time.sleep(TIMEOUTS['search_result'])

            
        except Exception as e:
            self.gui.log_message(f"✗ Erro ao buscar paciente: {str(e)}", "error")

    def create_patient(self):
        """Cria paciente"""
        try:
            self.gui.log_message("→ Criando paciente...", "info")

            anchor = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input#nomeNascimento + a.table-editable-ancora")
                )
            )
            anchor.click()

            civil_name = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.visibility_of_element_located((By.ID, SELECTORS['nome_nascimento']))
            )

            civil_name.clear()
            civil_name.send_keys(PATIENT_NAME) 

            self.gui.log_message(f"✓ Nome '{PATIENT_NAME}' inserido no campo", "success")

            anchor = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input#dataNascmentoDoPacienteFluxo + a.table-editable-ancora")
                )
            )
            anchor.click()

            data_nascimento = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.visibility_of_element_located((By.ID, SELECTORS['data_nascimento']))
            )

            data_nascimento.clear()
            data_nascimento.send_keys("29/10/2001") 

            self.gui.log_message(f"✓ Data de nascimento inserida no campo", "success")

            anchor = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input#telefone + a.table-editable-ancora")
                )
            )
            anchor.click()

            celular = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.visibility_of_element_located((By.ID, SELECTORS['celular']))
            )

            celular.clear()
            celular.send_keys("(43) 98406-5558") 

            self.gui.log_message(f"✓ Telefone inserido no campo", "success")


            # Adicionar Api do ViaCep e preencher campos do endereço

        except Exception as e:
            self.gui.log_message(f"✗ Erro ao criar paciente: {str(e)}", "error")
            
    def quit_driver(self):
        """Fecha o driver do Chrome"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
        self.root.destroy()
