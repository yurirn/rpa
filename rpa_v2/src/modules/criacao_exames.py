import time
import traceback
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from config import SELECTORS, TIMEOUTS, PATIENT_NAME
from src.utils.viacep_client import buscar_endereco

class ExamAutomation:
    def __init__(self):
        self.driver = None
        self.wait = None

    def setup(self, url):
        self.driver = BrowserFactory.create_chrome()
        self.wait = WebDriverWait(self.driver, TIMEOUTS['element_wait'])
        self.driver.get(url)

    def login(self, username, password):
        try:
            user_field = self.wait.until(EC.presence_of_element_located((By.ID, "username")))
            user_field.clear()
            user_field.send_keys(username)
            pass_field = self.driver.find_element(By.ID, "password")
            pass_field.clear()
            pass_field.send_keys(password)
            self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
            log_message("✓ Login realizado", "SUCCESS")
        except Exception as e:
            log_message(f"✗ Erro no login: {str(e)}", "ERROR")

    def is_initial_screen(self):
        try:
            btn = self.driver.find_element(By.XPATH, SELECTORS['create_exam_button'])
            return btn.is_displayed()
        except:
            return False

    def is_patient_screen(self):
        try:
            el = self.driver.find_element(By.CSS_SELECTOR, "input#nomeNascimento + a.table-editable-ancora")
            return el.is_displayed()
        except:
            return False

    def is_cadastro_paciente_screen(self):
        try:
            form = self.driver.find_element(By.ID, "cadastroPaciente")
            return form.is_displayed()
        except:
            return False

    def select_exam_type(self):
        try:
            log_message("→ Selecionando tipo de exame", "INFO")
            exam_select = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "form#formCriarNovoExame select")))
            self.driver.execute_script("""
                arguments[0].value = '175';
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, exam_select)
            log_message("✓ Tipo de exame selecionado: AN - Anátomo Patológico", "SUCCESS")
            time.sleep(1)
        except Exception as e:
            log_message(f"✗ Erro ao selecionar tipo de exame: {str(e)}", "ERROR")

    def search_patient(self):
        try:
            log_message("→ Buscando paciente...", "INFO")
            time.sleep(TIMEOUTS['page_load'])
            search_field = self.wait.until(EC.presence_of_element_located((By.ID, SELECTORS['patient_search_field'])))
            search_field.clear()
            search_field.send_keys(PATIENT_NAME)
            consult_button = self.wait.until(EC.presence_of_element_located((By.ID, SELECTORS['consult_button'])))
            self.driver.execute_script("arguments[0].click();", consult_button)
            time.sleep(TIMEOUTS['search_result'])
            if self.check_existing_patient():
                log_message("✓ Paciente encontrado - selecionando", "SUCCESS")
                return True
            else:
                log_message("→ Paciente não encontrado - criando novo", "INFO")
                create_patient = self.wait.until(EC.presence_of_element_located((By.ID, SELECTORS['create_patient'])))
                self.driver.execute_script("arguments[0].click();", create_patient)
                time.sleep(TIMEOUTS['search_result'])
                return False
        except Exception as e:
            log_message(f"✗ Erro ao buscar paciente: {str(e)}", "ERROR")
            return False

    def check_existing_patient(self):
        try:
            time.sleep(2)
            rows = self.driver.find_elements(By.CSS_SELECTOR, "#formPacienteId table tbody tr")
            if not rows:
                return False
            expected_birth_date = "29/10/2001"
            for row in rows:
                try:
                    name = row.find_element(By.CSS_SELECTOR, "td:nth-child(2) span").text.strip().upper()
                    birth = row.find_element(By.CSS_SELECTOR, "td:nth-child(5)").text.strip()
                    if (PATIENT_NAME.upper() in name and birth == expected_birth_date):
                        radio = row.find_element(By.CSS_SELECTOR, "input[type='radio'].pacienteId")
                        self.driver.execute_script("arguments[0].click();", radio)
                        patient_selected = self.wait.until(EC.presence_of_element_located((By.ID, SELECTORS['patient_selected'])))
                        self.driver.execute_script("arguments[0].click();", patient_selected)
                        return True
                except:
                    continue
            return False
        except:
            return False

    def create_patient(self):
        try:
            log_message("→ Criando paciente...", "INFO")
            anchor = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input#nomeNascimento + a.table-editable-ancora")))
            anchor.click()
            civil_name = self.wait.until(EC.visibility_of_element_located((By.ID, SELECTORS['nome_nascimento'])))
            civil_name.clear()
            civil_name.send_keys(PATIENT_NAME)
            anchor = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input#dataNascmentoDoPacienteFluxo + a.table-editable-ancora")))
            anchor.click()
            data_nascimento = self.wait.until(EC.visibility_of_element_located((By.ID, SELECTORS['data_nascimento'])))
            data_nascimento.clear()
            data_nascimento.send_keys("29/10/2001")
            anchor = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input#telefone + a.table-editable-ancora")))
            anchor.click()
            celular = self.wait.until(EC.visibility_of_element_located((By.ID, SELECTORS['celular'])))
            celular.clear()
            celular.send_keys("(43) 98406-5558")
            dados_endereco = buscar_endereco("PR", "Londrina", "Rua Cesar de Oliveira Bertin")
            for item in dados_endereco:
                anchor = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input#cep + a.table-editable-ancora")))
                anchor.click()
                cep = self.wait.until(EC.visibility_of_element_located((By.ID, SELECTORS['cep'])))
                cep.clear()
                cep.send_keys(dados_endereco[0]["cep"])
                self.driver.execute_script("arguments[0].blur();", cep)
            log_message("✓ Paciente criado", "SUCCESS")
        except Exception as e:
            log_message(f"✗ Erro ao criar paciente: {str(e)}", "ERROR")

    def fill_exam_data(self):
        try:
            next_button = self.wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a.btn.btn-sm.btn-primary.chamadaAjax.setupAjax[data-url='/paciente/saveAjax']")
            ))
            self.driver.execute_script("arguments[0].click();", next_button)
            time.sleep(TIMEOUTS['page_load'])

            anchor = self.wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input#convenioInput + a.table-editable-ancora")
            ))
            valor_anchor = anchor.text.strip()
            if valor_anchor == "Vazio":
                anchor.click()
                convenioInput = self.wait.until(EC.visibility_of_element_located((By.ID, 'convenioInput')))
                convenioInput.clear()
                convenioInput.send_keys("UNIMED (LONDRINA)")
                time.sleep(0.5)
                convenioInput.send_keys(Keys.ENTER)
                time.sleep(1)

            anchor = self.wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input#codigoUsuarioConvenio + a.table-editable-ancora")
            ))
            anchor.click()
            numero_da_carteira = self.wait.until(EC.visibility_of_element_located((By.ID, 'codigoUsuarioConvenio')))
            numero_da_carteira.clear()
            numero_da_carteira.send_keys("0050000004252740")
            time.sleep(0.5)

            self.fill_doctor_field()
            self.fill_origin_field()
            self.add_exam_material()
        except Exception as e:
            log_message(f"✗ Erro ao preencher dados do exame: {str(e)}", "ERROR")

    def fill_doctor_field(self):
        time.sleep(2)
        try:
            anchor_medico = None
            try:
                anchor_medico = self.driver.find_element(By.XPATH, "//input[@id='medicoRequisitanteInput']/following-sibling::a[contains(@class, 'table-editable-ancora')]")
            except:
                pass
            if not anchor_medico:
                try:
                    anchor_medico = self.driver.find_element(By.XPATH, "//td[input[@id='medicoRequisitanteInput']]//a[contains(@class, 'table-editable-ancora')]")
                except:
                    pass
            if not anchor_medico:
                anchor_medico = self.wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//input[@id='medicoRequisitanteInput']/..//a[contains(@class, 'table-editable-ancora')]")
                    )
                )
            self.driver.execute_script("arguments[0].click();", anchor_medico)
            medico_input = self.wait.until(EC.visibility_of_element_located((By.ID, 'medicoRequisitanteInput')))
            medico_input.clear()
            medico_input.send_keys("LEONARDO OBA")
            time.sleep(1.5)
            try:
                dropdown_option = WebDriverWait(self.driver, 3).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//ul[contains(@class, 'typeahead')]//a[contains(text(), 'LEONARDO OBA')]")
                    )
                )
                dropdown_option.click()
                medico_input.send_keys(Keys.ENTER)
            except:
                medico_input.send_keys(Keys.ENTER)
        except Exception as e:
            log_message(f"✗ Erro ao inserir médico requisitante: {str(e)}", "ERROR")

    def fill_origin_field(self):
        try:
            time.sleep(2)
            anchor = self.driver.find_element(By.XPATH, "//input[@id='procedenciaInput']/following-sibling::a[contains(@class, 'table-editable-ancora')]")
            self.driver.execute_script("arguments[0].click();", anchor)
            input_el = self.driver.find_element(By.CSS_SELECTOR, "#procedenciaInput")
            input_el.clear()
            input_el.send_keys("GASTROCLINICA")
            time.sleep(1)
            try:
                option = self.driver.find_element(By.XPATH, "//ul[contains(@class, 'typeahead')]//a[contains(text(), 'GASTROCLINICA')]")
                option.click()
            except:
                input_el.send_keys(Keys.ENTER)
        except Exception as e:
            log_message(f"✗ Erro ao inserir procedência: {str(e)}", "ERROR")

    def add_exam_material(self):
        try:
            time.sleep(2)
            novo_material_link = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[@title='Novo material' and contains(@class, 'chamadaAjax')]")
                )
            )
            self.driver.execute_script("arguments[0].click();", novo_material_link)
            time.sleep(2)
            anchor_quantidade = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//input[@name='quantidadeRecipiente']/following-sibling::a[contains(@class, 'table-editable-ancora')]")
                )
            )
            self.driver.execute_script("arguments[0].click();", anchor_quantidade)
            time.sleep(1)
            quantidade_input = self.wait.until(EC.visibility_of_element_located((By.NAME, 'quantidadeRecipiente')))
            quantidade_input.clear()
            quantidade_input.send_keys("4")
        except Exception as e:
            log_message(f"✗ Erro ao adicionar material: {str(e)}", "ERROR")

    def finalize_exam_creation(self):
        try:
            time.sleep(2)
            selectors = [
                "//a[@data-url='/moduloExame/saveExameAjax']",
                "//a[@title='Próximo']",
                "//a[contains(@class, 'btn') and contains(text(), 'Próximo')]",
                "//a[contains(@class, 'chamadaAjax') and contains(@class, 'btn-primary')]",
                "//a[contains(text(), 'Próximo')]"
            ]
            for sel in selectors:
                try:
                    btn = self.driver.find_element(By.XPATH, sel)
                    btn.click()
                    log_message("✓ Botão 'Próximo' clicado - finalizando exame", "SUCCESS")
                    time.sleep(3)
                    return
                except:
                    continue
            log_message("✗ Botão 'Próximo' não encontrado", "ERROR")
        except Exception as e:
            log_message(f"✗ Erro ao finalizar exame: {str(e)}", "ERROR")

    def run(self, params):
        username = params.get("username")
        password = params.get("password")
        url = params.get("url", "https://pathoweb.com.br/login/auth")
        self.setup(url)
        try:
            log_message("Iniciando automação de criação de exame...", "INFO")
            self.login(username, password)
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=1']"))).click()
            time.sleep(2)
            if self.is_initial_screen():
                log_message("Na tela inicial - navegando para criação de exame", "INFO")
                self.driver.find_element(By.XPATH, SELECTORS['create_exam_button']).click()
                time.sleep(TIMEOUTS['page_load'])
                self.select_exam_type()
                modal_button = self.wait.until(EC.presence_of_element_located((By.XPATH, SELECTORS['modal_create_button'])))
                self.driver.execute_script("arguments[0].click();", modal_button)
            else:
                log_message("Já está na tela de exame", "SUCCESS")
            if not self.is_patient_screen():
                if not self.is_cadastro_paciente_screen():
                    found = self.search_patient()
                    if not found:
                        self.create_patient()
            else:
                log_message("Usando paciente existente selecionado", "SUCCESS")
            time.sleep(2)
            self.fill_exam_data()
            self.finalize_exam_creation()
            log_message("✓ Processo de criação de exame concluído!", "SUCCESS")
        except Exception as e:
            tb = traceback.format_exc()
            log_message(f"✗ Erro no processo de criação de exame: {str(e)}\n{tb}", "ERROR")
        finally:
            self.driver.quit()

def run(params):
    automation = ExamAutomation()
    automation.run(params)