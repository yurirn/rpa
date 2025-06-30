from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
from config import CHROME_OPTIONS, SELECTORS, LOGIN_SUCCESS_PATTERNS, LOGIN_FAIL_PATTERNS, TIMEOUTS, PATIENT_NAME
from viacep_client import buscar_endereco
import traceback

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
            
            user_field = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.presence_of_element_located((By.ID, SELECTORS['username_field']))
            )
            pass_field = self.driver.find_element(By.ID, SELECTORS['password_field'])
            
            user_field.clear()
            user_field.send_keys(username)
            time.sleep(1)
            pass_field.clear()
            pass_field.send_keys(password)
            
            login_button = self.driver.find_element(By.XPATH, SELECTORS['login_button'])
            login_button.click()
            
            self.gui.log_message("✓ Credenciais enviadas", "success")
            
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
        
    def is_initial_screen(self):
        """Verifica se está na tela inicial onde tem o botão de criar exame"""
        try:
            create_button = self.driver.find_element(By.XPATH, SELECTORS['create_exam_button'])
            return create_button.is_displayed()
        except:
            return False
        
    def is_patient_screen(self):
        """Verifica se o paciente já esta selecionado"""
        try:
            patient_screen = self.driver.find_element(By.CSS_SELECTOR, "input#nomeNascimento + a.table-editable-ancora")
            return patient_screen.is_displayed()
        except:
            return False
            
    def create_new_exam(self):
        """Processo completo de criação de exame"""
        try:
            self.gui.log_message("→ Iniciando criação de exame", "info")
            
            if self.is_initial_screen():
                self.gui.log_message("✓ Na tela inicial - clicando nos botões", "info")
            
                first_button = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                    EC.element_to_be_clickable((By.XPATH, SELECTORS['create_exam_button']))
                )
                first_button.click()
                self.gui.log_message("✓ Primeiro botão clicado", "success")
                
                time.sleep(TIMEOUTS['page_load'])
                
                self.select_exam_type()
                
                modal_button = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                   EC.presence_of_element_located((By.XPATH, SELECTORS['modal_create_button']))
                )
                self.driver.execute_script("arguments[0].click();", modal_button)
                self.gui.log_message("✓ Botão do modal clicado", "success")
            else:
                self.gui.log_message("✓ Já está na tela de exame - pulando botões", "success")
            
            if not self.is_patient_screen():    
                patient_found = self.search_patient()
                self.gui.log_message(patient_found, "success")
                if not patient_found:
                    self.create_patient()
            else:
                self.gui.log_message("✓ Usando paciente existente selecionado", "success")
            
            time.sleep(2) 
            
            next_button = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "a.btn.btn-sm.btn-primary.chamadaAjax.setupAjax[data-url='/paciente/saveAjax']")
                )
            )
            self.driver.execute_script("arguments[0].click();", next_button)
            self.gui.log_message("✓ Botão 'Próximo' clicado para salvar paciente", "success")
            
            time.sleep(TIMEOUTS['page_load']) 

            anchor = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input#convenioInput + a.table-editable-ancora")
                )
            )

            valor_anchor = anchor.text.strip()
            if valor_anchor == "Vazio":
                anchor.click()

                convenioInput = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                    EC.visibility_of_element_located((By.ID, 'convenioInput'))
                )

                convenioInput.clear()
                convenioInput.send_keys("UNIMED (LONDRINA)") 
                time.sleep(0.5)
                convenioInput.send_keys(Keys.ENTER)   
                self.gui.log_message(f"✓ Convênio informado", "success")

            anchor = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input#codigoUsuarioConvenio + a.table-editable-ancora")
                )
            )
            anchor.click()

            numero_da_carteira = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.visibility_of_element_located((By.ID, 'codigoUsuarioConvenio'))
            )

            numero_da_carteira.clear()
            numero_da_carteira.send_keys("0050000004252740") 

            self.gui.log_message(f"✓ Numero da carteirinha informado", "success")

            # Inserir médico requisitante
            time.sleep(2)  # Aguardar carregamento da página
            
            try:
                # Tentar encontrar a âncora do médico requisitante de diferentes formas
                anchor_medico = None
                
                # Primeira tentativa: âncora próxima ao input
                try:
                    anchor_medico = self.driver.find_element(By.XPATH, "//input[@id='medicoRequisitanteInput']/following-sibling::a[contains(@class, 'table-editable-ancora')]")
                except:
                    pass
                
                # Segunda tentativa: âncora com texto "Vazio" na mesma célula
                if not anchor_medico:
                    try:
                        anchor_medico = self.driver.find_element(By.XPATH, "//td[input[@id='medicoRequisitanteInput']]//a[contains(@class, 'table-editable-ancora')]")
                    except:
                        pass
                
                # Terceira tentativa: qualquer âncora table-editable próxima ao input do médico
                if not anchor_medico:
                    anchor_medico = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//input[@id='medicoRequisitanteInput']/..//a[contains(@class, 'table-editable-ancora')]")
                        )
                    )
                
                self.driver.execute_script("arguments[0].click();", anchor_medico)
                self.gui.log_message("✓ Campo médico ativado", "success")

                medico_input = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                    EC.visibility_of_element_located((By.ID, 'medicoRequisitanteInput'))
                )

                medico_input.clear()
                medico_input.send_keys("LEONARDO OBA")
                
                # Aguardar dropdown aparecer e selecionar opção
                time.sleep(1.5)
                
                try:
                    # Procurar pela opção que contém "LEONARDO OBA" no dropdown
                    dropdown_option = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//ul[contains(@class, 'typeahead')]//a[contains(text(), 'LEONARDO OBA')]")
                        )
                    )
                    dropdown_option.click()
                    self.gui.log_message(f"✓ Médico 'LEONARDO OBA' selecionado do dropdown", "success")
                    medico_input.send_keys(Keys.ENTER)
                except:
                    # Se não encontrar no dropdown, tentar pressionar Enter
                    medico_input.send_keys(Keys.ENTER)
                    self.gui.log_message(f"✓ Nome do médico 'LEONARDO OBA' inserido", "success")
                    
            except Exception as e:
                self.gui.log_message(f"✗ Erro ao inserir médico requisitante: {str(e)}", "error")
                # Continuar execução sem parar por este erro

            # Inserir procedência
            time.sleep(2)  # Aguardar carregamento após seleção do médico
            
            try:
                self.gui.log_message("→ Procurando âncora da procedência...", "info")
                
                # Primeiro clicar na âncora para ativar o campo
                anchor_procedencia = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//input[@id='procedenciaInput']/following-sibling::a[contains(@class, 'table-editable-ancora')]")
                    )
                )
                self.driver.execute_script("arguments[0].click();", anchor_procedencia)
                self.gui.log_message("✓ Âncora da procedência clicada - campo ativado", "success")

                # Agora encontrar o campo de input que deve estar visível
                procedencia_input = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                    EC.visibility_of_element_located((By.ID, 'procedenciaInput'))
                )
                self.gui.log_message("✓ Campo de procedência encontrado e visível", "success")

                self.gui.log_message("→ Preenchendo campo de procedência...", "info")
                procedencia_input.clear()
                procedencia_input.send_keys("GASTROCLINICA")
                self.gui.log_message("✓ Texto 'GASTROCLINICA' inserido no campo", "success")
                
                # Aguardar dropdown aparecer e selecionar opção
                time.sleep(1.5)
                
                try:
                    self.gui.log_message("→ Procurando opção no dropdown...", "info")
                    # Procurar pela opção "GASTROCLINICA" no dropdown
                    dropdown_option = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//ul[contains(@class, 'typeahead')]//a[contains(text(), 'GASTROCLINICA')]")
                        )
                    )
                    dropdown_option.click()
                    self.gui.log_message(f"✓ Procedência 'GASTROCLINICA' selecionada do dropdown", "success")
                except Exception as dropdown_error:
                    self.gui.log_message(f"→ Dropdown não encontrado ({str(dropdown_error)}), tentando Enter...", "warning")
                    # Se não encontrar no dropdown, tentar pressionar Enter
                    procedencia_input.send_keys(Keys.ENTER)
                    self.gui.log_message(f"✓ Procedência 'GASTROCLINICA' inserida com Enter", "success")
                    
            except Exception as e:
                tb = traceback.format_exc()
                self.gui.log_message(f"✗ Erro ao inserir procedência: {str(e)}\nTraceback: {tb}", "error")
                # Continuar execução sem parar por este erro

            # Adicionar novo material
            time.sleep(2)  # Aguardar carregamento após procedência
            
            try:
                self.gui.log_message("→ Procurando botão 'Novo material'...", "info")
                
                # Encontrar e clicar no link "Novo material"
                novo_material_link = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//a[@title='Novo material' and contains(@class, 'chamadaAjax')]")
                    )
                )
                self.driver.execute_script("arguments[0].click();", novo_material_link)
                self.gui.log_message("✓ Botão 'Novo material' clicado", "success")
                
                # Aguardar carregar o novo campo
                time.sleep(2)
                
                # Encontrar e clicar na âncora da quantidade de recipiente
                self.gui.log_message("→ Procurando campo de quantidade de recipiente...", "info")
                anchor_quantidade = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//input[@name='quantidadeRecipiente']/following-sibling::a[contains(@class, 'table-editable-ancora')]")
                    )
                )
                self.driver.execute_script("arguments[0].click();", anchor_quantidade)
                self.gui.log_message("✓ Âncora da quantidade clicada - campo ativado", "success")
                
                # Encontrar o campo de quantidade e alterar para 4
                quantidade_input = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                    EC.visibility_of_element_located((By.NAME, 'quantidadeRecipiente'))
                )
                
                quantidade_input.clear()
                quantidade_input.send_keys("4")
                self.gui.log_message("✓ Quantidade alterada para 4", "success")
                
            except Exception as e:
                tb = traceback.format_exc()
                self.gui.log_message(f"✗ Erro ao adicionar material: {str(e)}\nTraceback: {tb}", "error")
                # Continuar execução sem parar por este erro

            # Clicar no botão Próximo para finalizar
            time.sleep(2)  # Aguardar carregamento após adicionar material
            
            try:
                self.gui.log_message("→ Procurando botão 'Próximo'...", "info")
                
                # Múltiplas tentativas para encontrar o botão "Próximo"
                proximo_button = None
                
                # Tentativa 1: Por data-url específico
                try:
                    proximo_button = self.driver.find_element(By.XPATH, "//a[@data-url='/moduloExame/saveExameAjax']")
                    self.gui.log_message("✓ Botão 'Próximo' encontrado por data-url", "success")
                except:
                    pass
                
                # Tentativa 2: Por title="Próximo"
                if not proximo_button:
                    try:
                        proximo_button = self.driver.find_element(By.XPATH, "//a[@title='Próximo']")
                        self.gui.log_message("✓ Botão 'Próximo' encontrado por title", "success")
                    except:
                        pass
                
                # Tentativa 3: Por texto "Próximo" com classes btn
                if not proximo_button:
                    try:
                        proximo_button = self.driver.find_element(By.XPATH, "//a[contains(@class, 'btn') and contains(text(), 'Próximo')]")
                        self.gui.log_message("✓ Botão 'Próximo' encontrado por texto e classe", "success")
                    except:
                        pass
                
                # Tentativa 4: Por classes específicas
                if not proximo_button:
                    try:
                        proximo_button = self.driver.find_element(By.XPATH, "//a[contains(@class, 'chamadaAjax') and contains(@class, 'btn-primary')]")
                        self.gui.log_message("✓ Botão 'Próximo' encontrado por classes", "success")
                    except:
                        pass
                
                # Tentativa 5: Qualquer link com "Próximo"
                if not proximo_button:
                    try:
                        proximo_button = self.driver.find_element(By.XPATH, "//a[contains(text(), 'Próximo')]")
                        self.gui.log_message("✓ Botão 'Próximo' encontrado por texto", "success")
                    except:
                        pass
                
                if proximo_button:
                    # Aguardar até o botão estar clicável
                    WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable(proximo_button))
                    self.driver.execute_script("arguments[0].click();", proximo_button)
                    self.gui.log_message("✓ Botão 'Próximo' clicado - finalizando exame", "success")
                    
                    # Aguardar processamento
                    time.sleep(3)
                    self.gui.log_message("✓ Processo de criação de exame finalizado!", "success")
                else:
                    self.gui.log_message("✗ Botão 'Próximo' não encontrado em nenhuma tentativa", "error")
                    # Tentar listar todos os botões disponíveis para debug
                    try:
                        buttons = self.driver.find_elements(By.TAG_NAME, "a")
                        button_texts = [btn.text.strip() for btn in buttons if btn.text.strip()]
                        self.gui.log_message(f"→ Botões disponíveis: {button_texts[:10]}", "info")  # Mostrar só os primeiros 10
                    except:
                        pass
                
            except Exception as e:
                tb = traceback.format_exc()
                self.gui.log_message(f"✗ Erro ao clicar no botão Próximo: {str(e)}\nTraceback: {tb}", "error")
                # Continuar execução sem parar por este erro
            
        except Exception as e:         
            tb = traceback.format_exc()
            self.gui.log_message(f"✗ Erro ao criar exame: {str(e)}\n{tb}", "error")

    def select_exam_type(self):
        """Seleciona o tipo de exame no modal"""
        try:
            self.gui.log_message("→ Selecionando tipo de exame", "info")
        
            time.sleep(3)
            
            exam_select_element = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "form#formCriarNovoExame select"))
            )
            
            self.driver.execute_script("""
                arguments[0].value = '175';
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, exam_select_element)
            
            self.gui.log_message("✓ Tipo de exame selecionado: AN - Anátomo Patológico", "success")
            time.sleep(1)
        
        except Exception as e:
            tb = traceback.format_exc()
            self.gui.log_message(f"✗ Erro ao selecionar tipo de exame: {str(e)}\n{tb}", "error")

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
            
            time.sleep(TIMEOUTS['search_result']) 
            self.gui.log_message("✓ Busca de paciente realizada", "success")

            if self.check_existing_patient():
                self.gui.log_message("✓ Paciente encontrado na tabela - selecionando", "success")
                return True
            else:
                self.gui.log_message("→ Paciente não encontrado - criando novo", "info")
                
                create_patient = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                    EC.presence_of_element_located((By.ID, SELECTORS['create_patient']))
                )

                time.sleep(TIMEOUTS['search_result']) 
                self.driver.execute_script("arguments[0].click();", create_patient)
                self.gui.log_message("✓ Botão criar paciente clicado via Javascript", "success")

                time.sleep(TIMEOUTS['search_result'])
                return False
            
        except Exception as e:
            tb = traceback.format_exc()
            self.gui.log_message(f"✗ Erro ao buscar paciente: {str(e)}\n{tb}", "error")
            return False

    def check_existing_patient(self):
        """Verifica se o paciente já existe na tabela de resultados"""
        try:
            self.gui.log_message("→ Verificando se paciente já existe na tabela...", "info")
        
            time.sleep(2)
            
            table_rows = self.driver.find_elements(By.CSS_SELECTOR, "#formPacienteId table tbody tr")
            self.gui.log_message(table_rows)
            
            if not table_rows:
                self.gui.log_message("→ Nenhuma linha encontrada na tabela", "info")
                return False
            
            expected_birth_date = "10/10/1990" 
            
            for row in table_rows:
                try:
                    name_element = row.find_element(By.CSS_SELECTOR, "td:nth-child(2) span")
                    patient_name = name_element.text.strip().upper()
                    self.gui.log_message(patient_name)

                    birth_date_element = row.find_element(By.CSS_SELECTOR, "td:nth-child(5)")
                    birth_date = birth_date_element.text.strip()
                    self.gui.log_message(birth_date)
                    
                    self.gui.log_message(f"→ Verificando: {patient_name} - {birth_date}", "info")
                    
                    if (PATIENT_NAME.upper() in patient_name and birth_date == expected_birth_date):
                        
                        radio_button = row.find_element(By.CSS_SELECTOR, "input[type='radio'].pacienteId")
                        self.driver.execute_script("arguments[0].click();", radio_button)

                        self.gui.log_message(radio_button)
                        
                        patient_id = radio_button.get_attribute("value")
                        self.gui.log_message(f"✓ Paciente encontrado e selecionado - ID: {patient_id}", "success")
                        self.gui.log_message(f"✓ Nome: {patient_name} | Data: {birth_date}", "success")

                        patient_selected = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                            EC.presence_of_element_located((By.ID, SELECTORS['patient_selected']))
                        )
                        self.driver.execute_script("arguments[0].click();", patient_selected)
                        self.gui.log_message("✓ Botão Utilizar paciente selecionado clicado via JavaScript", "success")
                        
                        return True
                        
                except Exception as e:
                    continue
            
            self.gui.log_message("→ Paciente não encontrado na tabela de resultados", "info")
            return False
            
        except Exception as e:
            tb = traceback.format_exc()
            self.gui.log_message(f"✗ Erro ao verificar paciente existente: {str(e)}\n{tb}", "error")
            return False

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

            dados_endereco = buscar_endereco("PR", "Londrina", "Rua Cesar de Oliveira Bertin")
            for item in dados_endereco:
                anchor = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "input#cep + a.table-editable-ancora")
                    )
                )
                anchor.click()

                cep = WebDriverWait(self.driver, TIMEOUTS['element_wait']).until(
                    EC.visibility_of_element_located((By.ID, SELECTORS['cep']))
                )

                cep.clear()
                cep.send_keys(dados_endereco[0]["cep"]) 

                self.driver.execute_script("arguments[0].blur();", cep)

            self.gui.log_message(f"✓ Cep informado", "success")                     

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
