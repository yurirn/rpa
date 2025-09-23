import json
import time
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

# Importar funções do OCR
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from exame_data_extractor import process_all_files_optimized

class CriacaoExamesHclSus(BaseModule):
    def __init__(self):
        super().__init__(nome="Criação Exame Hospital Câncer")
        self.driver = None
        self.wait = None
        self.dados_exame = {}

    def carregar_dados_ocr(self, arquivo_json):
        """Carrega os dados extraídos pelo OCR"""
        try:
            with open(arquivo_json, 'r', encoding='utf-8') as f:
                dados = json.load(f)
                self.dados_exame = dados.get('dados_exame', {})
                nome_paciente = self.dados_exame.get('paciente', 'Não encontrado') if self.dados_exame else 'Não encontrado'
                log_message(f"✅ Dados carregados: {nome_paciente}", "SUCCESS")
                return True
        except Exception as e:
            log_message(f"❌ Erro ao carregar dados: {str(e)}", "ERROR")
            return False

    def setup(self, url):
        self.driver = BrowserFactory.create_chrome()
        self.wait = WebDriverWait(self.driver, 10)
        self.driver.get(url)

    def login(self, username, password):
        """Mesmo login do módulo original"""
        try:
            user_field = self.wait.until(EC.presence_of_element_located((By.ID, "username")))
            user_field.clear()
            user_field.send_keys(username)
            pass_field = self.driver.find_element(By.ID, "password")
            pass_field.clear()
            pass_field.send_keys(password)
            self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
            log_message("✔ Login realizado", "SUCCESS")
        except Exception as e:
            log_message(f"✗ Erro no login: {str(e)}", "ERROR")

    def select_exam_type(self):
        """Seleciona tipo de exame baseado no convênio"""
        try:
            exam_type = "171"

            exam_select = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "form#formCriarNovoExame select"))
            )
            self.driver.execute_script(f"""
                arguments[0].value = '{exam_type}';
                arguments[0].dispatchEvent(new Event('change', {{ bubbles: true }}));
            """, exam_select)

            log_message(f"✔ Tipo de exame selecionado: {exam_type}", "SUCCESS")
            time.sleep(1)
        except Exception as e:
            log_message(f"✗ Erro ao selecionar tipo: {str(e)}", "ERROR")

    def search_or_create_patient(self):
        """Busca ou cria paciente usando dados do OCR"""
        try:
            nome_paciente = self.dados_exame.get('paciente', '')
            nascimento = self.dados_exame.get('nascimento', '')

            if not nome_paciente:
                log_message("❌ Nome do paciente não encontrado nos dados", "ERROR")
                return False

            # Buscar paciente
            search_field = self.wait.until(
                EC.presence_of_element_located((By.ID, "pacienteSearch"))
            )
            search_field.clear()
            search_field.send_keys(nome_paciente)

            consult_button = self.wait.until(
                EC.presence_of_element_located((By.ID, "consultarPaciente"))
            )
            self.driver.execute_script("arguments[0].click();", consult_button)
            time.sleep(3)

            # Verificar se encontrou
            rows = self.driver.find_elements(By.CSS_SELECTOR, "#formPacienteId table tbody tr")

            if rows and len(rows) > 0:
                # Selecionar o primeiro resultado
                radio = rows[0].find_element(By.CSS_SELECTOR, "input[type='radio'].pacienteId")
                self.driver.execute_script("arguments[0].click();", radio)

                patient_selected = self.wait.until(
                    EC.presence_of_element_located((By.ID, "usarPacienteSelecionado"))
                )
                self.driver.execute_script("arguments[0].click();", patient_selected)
                log_message(f"✔ Paciente encontrado: {nome_paciente}", "SUCCESS")
                return True
            else:
                # Criar novo paciente
                log_message(f"→ Criando novo paciente: {nome_paciente}", "INFO")
                create_patient = self.wait.until(
                    EC.presence_of_element_located((By.ID, "criarPaciente"))
                )
                self.driver.execute_script("arguments[0].click();", create_patient)
                time.sleep(2)

                # Preencher dados do novo paciente
                self.fill_new_patient_data()
                return True

        except Exception as e:
            log_message(f"✗ Erro ao buscar/criar paciente: {str(e)}", "ERROR")
            return False

    def fill_new_patient_data(self):
        """Preenche dados do novo paciente usando dados do OCR"""
        try:
            # Nome
            anchor = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input#nomeNascimento + a.table-editable-ancora"))
            )
            anchor.click()
            civil_name = self.wait.until(EC.visibility_of_element_located((By.ID, "nomeNascimento")))
            civil_name.clear()
            civil_name.send_keys(self.dados_exame.get('paciente', ''))

            # Data nascimento
            if self.dados_exame.get('nascimento'):
                anchor = self.wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "input#dataNascmentoDoPacienteFluxo + a.table-editable-ancora"))
                )
                anchor.click()
                data_nascimento = self.wait.until(
                    EC.visibility_of_element_located((By.ID, "dataNascmentoDoPacienteFluxo"))
                )
                data_nascimento.clear()
                data_nascimento.send_keys(self.dados_exame['nascimento'])

            log_message("✔ Dados do paciente preenchidos", "SUCCESS")

        except Exception as e:
            log_message(f"✗ Erro ao preencher dados: {str(e)}", "ERROR")

    def fill_exam_data(self):
        """Preenche dados do exame usando informações do OCR"""
        try:
            # Avançar para próxima tela
            next_button = self.wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "a.btn.btn-sm.btn-primary.chamadaAjax.setupAjax[data-url='/paciente/saveAjax']")
                )
            )
            self.driver.execute_script("arguments[0].click();", next_button)
            time.sleep(2)

            # Definir convênio baseado nos dados
            convenio = self.dados_exame.get('convenio', 'SUS')

            anchor = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input#convenioInput + a.table-editable-ancora"))
            )

            if anchor.text.strip() == "Vazio":
                anchor.click()
                convenioInput = self.wait.until(
                    EC.visibility_of_element_located((By.ID, 'convenioInput'))
                )
                convenioInput.clear()

                # Mapear convênio
                if 'SUS' in convenio.upper():
                    convenioInput.send_keys("SUS")
                else:
                    convenioInput.send_keys(convenio)

                time.sleep(0.5)
                convenioInput.send_keys(Keys.ENTER)
                time.sleep(1)

            # Preencher prontuário se houver
            if self.dados_exame.get('prontuario'):
                self.preencher_prontuario(self.dados_exame['prontuario'])

            # Preencher código de controle se houver
            if self.dados_exame.get('atendimento'):
                self.preencher_codigo_controle(self.dados_exame['atendimento'])

            # Preencher médico se houver
            if self.dados_exame.get('medico'):
                self.preencher_medico_requisitante(self.dados_exame['medico'])

            # Procedência padrão para Hospital do Câncer
            self.preencher_procedencia("HOSPITAL DO CANCER DE LONDRINA")

            self.preencher_procedimento()

            log_message("✔ Dados do exame preenchidos", "SUCCESS")

        except Exception as e:
            log_message(f"✗ Erro ao preencher exame: {str(e)}", "ERROR")

    def preencher_procedimento(self):
        """Preenche o procedimento padrão para exames do Hospital do Câncer"""
        try:
            procedimento = "020302004"
            log_message("Iniciando preenchimento do procedimento...", "INFO")

            # Passo 1: Clicar na âncora "Vazio"
            log_message("Passo 1: Procurando âncora 'Vazio' do procedimento...", "INFO")
            anchor = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH,
                     "//tr[@id='novosProcedimentos']//a[contains(@class, 'table-editable-ancora') and text()='Vazio']")
                )
            )
            self.driver.execute_script("arguments[0].click();", anchor)
            time.sleep(2)
            log_message("Âncora 'Vazio' clicada", "SUCCESS")

            # Passo 2: Aguardar input aparecer
            log_message("Passo 2: Aguardando input aparecer...", "INFO")
            procedimento_input = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//tr[@id='novosProcedimentos']//input[@id='procedimentoInput_novo']"))
            )
            time.sleep(1)  # Aguardar estabilizar
            log_message("Input do procedimento encontrado", "SUCCESS")

            # Passo 3: Inserir valor com JavaScript (único método que funciona)
            log_message(f"Passo 3: Inserindo '{procedimento}' com JavaScript...", "INFO")
            self.driver.execute_script("""
                var element = arguments[0];
                var value = arguments[1];
                element.focus();
                element.value = '';
                element.value = value;
                element.dispatchEvent(new Event('input', { bubbles: true }));
                element.dispatchEvent(new Event('change', { bubbles: true }));
            """, procedimento_input, procedimento)
            time.sleep(1)

            # Verificar se inseriu
            valor_inserido = procedimento_input.get_attribute("value")
            log_message(f"Valor inserido: '{valor_inserido}'", "INFO")

            # Passo 4: Dar ENTER
            log_message("Passo 4: Enviando ENTER...", "INFO")
            procedimento_input.send_keys(Keys.ENTER)
            time.sleep(3)
            log_message("ENTER enviado", "SUCCESS")

            log_message("Procedimento preenchido com sucesso!", "SUCCESS")
            return True

        except Exception as e:
            log_message(f"Erro ao inserir procedimento: {str(e)}", "ERROR")
            return False

    def preencher_prontuario(self, prontuario):
        try:
            anchor = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input#prontuario + a.table-editable-ancora"))
            )

            if anchor.text.strip() == "Vazio":
                anchor.click()
                prontuarioInput = self.wait.until(
                    EC.visibility_of_element_located((By.ID, 'prontuario'))
                )
                prontuarioInput.clear()

                prontuarioInput.send_keys(prontuario)

                time.sleep(0.5)
                prontuarioInput.send_keys(Keys.ENTER)
                time.sleep(1)

            log_message(f"✔ Prontuário preenchido: {prontuario}", "SUCCESS")

        except Exception as e:
            log_message(f"✗ Erro ao inserir prontuário: {str(e)}", "ERROR")

    def preencher_codigo_controle(self, atendimento):
        try:
            anchor = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input#codigoGuiaConvenio + a.table-editable-ancora"))
            )

            if anchor.text.strip() == "Vazio":
                anchor.click()
                codigo_controle = self.wait.until(
                    EC.visibility_of_element_located((By.ID, 'codigoGuiaConvenio'))
                )
                codigo_controle.clear()

                codigo_controle.send_keys(atendimento)

                time.sleep(0.5)
                codigo_controle.send_keys(Keys.ENTER)
                time.sleep(1)

            log_message(f"✔ Código de controle preenchido: {atendimento}", "SUCCESS")

        except Exception as e:
            log_message(f"✗ Erro ao inserir código de controle: {str(e)}", "ERROR")

    def preencher_medico_requisitante(self, nome_medico):
        """Preenche campo do médico"""
        try:
            anchor = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH,
                     "//input[@id='medicoRequisitanteInput']/following-sibling::a[contains(@class, 'table-editable-ancora')]")
                )
            )
            self.driver.execute_script("arguments[0].click();", anchor)

            medico_input = self.wait.until(
                EC.visibility_of_element_located((By.ID, 'medicoRequisitanteInput'))
            )
            #medico_input.clear()
            medico_input.send_keys(nome_medico)
            time.sleep(1)
            medico_input.send_keys(Keys.TAB)
            time.sleep(1)

            log_message(f"✔ Médico preenchido: {nome_medico}", "SUCCESS")

        except Exception as e:
            log_message(f"✗ Erro ao inserir médico: {str(e)}", "ERROR")

    def preencher_procedencia(self, procedencia):
        """Preenche campo de procedência"""
        try:
            anchor = self.driver.find_element(
                By.XPATH,
                "//input[@id='procedenciaInput']/following-sibling::a[contains(@class, 'table-editable-ancora')]"
            )
            self.driver.execute_script("arguments[0].click();", anchor)

            input_el = self.driver.find_element(By.CSS_SELECTOR, "#procedenciaInput")
            #input_el.clear()
            input_el.send_keys(procedencia)
            time.sleep(1)
            input_el.send_keys(Keys.TAB)
            time.sleep(1)

            log_message(f"✔ Procedência preenchida: {procedencia}", "SUCCESS")

        except Exception as e:
            log_message(f"✗ Erro ao inserir procedência: {str(e)}", "ERROR")

    def processar_um_exame(self, arquivo_info):
        """Processa um único exame com base nas informações do OCR"""
        try:
            nome_arquivo = arquivo_info['arquivo_origem']
            log_message("")
            base_name = os.path.splitext(nome_arquivo)[0]
            json_path = os.path.join('resultados_ocr', f"{base_name}_dados.json")

            log_message(f"🔄 Processando exame: {nome_arquivo}", "INFO")

            # Carregar dados do OCR
            if not self.carregar_dados_ocr(json_path):
                log_message(f"❌ Falha ao carregar dados do OCR para {nome_arquivo}", "ERROR")
                return False

            # Navegar para módulo (se não estiver já)
            try:
                self.wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=1']"))
                ).click()
                time.sleep(2)
            except:
                # Já está no módulo correto
                pass

            # Criar novo exame
            self.driver.find_element(By.XPATH, "//a[contains(text(), 'Criar novo exame')]").click()
            time.sleep(2)

            # Selecionar tipo
            #self.select_exam_type()

            # Clicar em criar
            modal_button = self.wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[@id='modalCriarNovoExame']//a[contains(text(), 'Criar novo exame')]")
                )
            )
            self.driver.execute_script("arguments[0].click();", modal_button)
            time.sleep(2)

            # Buscar ou criar paciente
            if not self.search_or_create_patient():
                log_message(f"❌ Falha ao processar paciente para {nome_arquivo}", "ERROR")
                return False

            # Preencher dados do exame
            self.fill_exam_data()

            btn = self.driver.find_element(By.XPATH, "//a[@data-url='/moduloExame/saveExameAjax']")
            btn.click()
            time.sleep(1.5)

            self.preencher_mascara()

            log_message(f"✔ Exame criado com sucesso para {nome_arquivo}!", "SUCCESS")
            return True

        except Exception as e:
            log_message(f"✗ Erro ao processar exame {nome_arquivo}: {str(e)}", "ERROR")
            return False

    def preencher_mascara(self):
        """Preenche a máscara escrevendo diretamente no iframe do CKEditor e fecha o exame"""
        try:
            texto_mascara = "Estudo Imuno-histoquímico."
            log_message("Iniciando preenchimento da máscara no iframe...", "INFO")

            # Passo 1: Encontrar o iframe do CKEditor
            log_message("Passo 1: Procurando iframe do CKEditor...", "INFO")
            iframe = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "iframe.cke_wysiwyg_frame"))
            )
            log_message("Iframe do CKEditor encontrado", "SUCCESS")

            # Passo 2: Mudar para o contexto do iframe
            log_message("Passo 2: Mudando para contexto do iframe...", "INFO")
            self.driver.switch_to.frame(iframe)

            # Passo 3: Encontrar o body dentro do iframe
            log_message("Passo 3: Procurando body dentro do iframe...", "INFO")
            body = self.wait.until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            log_message("Body encontrado dentro do iframe", "SUCCESS")

            # Debug: Verificar estado atual do body
            conteudo_atual = body.text
            log_message(f"Conteúdo atual do body: '{conteudo_atual}'", "INFO")

            # Passo 4: Adicionar o texto no body
            log_message(f"Passo 4: Inserindo texto '{texto_mascara}'...", "INFO")
            self.driver.execute_script("arguments[0].innerHTML = arguments[1];", body, texto_mascara)
            time.sleep(1)

            # Verificar se inseriu
            conteudo_apos_insercao = body.text
            log_message(f"Conteúdo após inserção: '{conteudo_apos_insercao}'", "INFO")

            # Passo 5: Voltar para o contexto principal
            log_message("Passo 5: Voltando para contexto principal...", "INFO")
            self.driver.switch_to.default_content()

            # Passo 6: Salvar as alterações
            log_message("Passo 6: Salvando alterações...", "INFO")
            try:
                btn_salvar = self.wait.until(
                    EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Salvar"))
                )
            except:
                btn_salvar = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[@data-url='/moduloExame/saveExameAjax']"))
                )

            self.driver.execute_script("arguments[0].click();", btn_salvar)
            time.sleep(3)  # Aguardar salvamento
            log_message("Alterações salvas", "SUCCESS")

            # Passo 7: Clicar no botão "Fechar exame"
            log_message("Passo 7: Procurando botão 'Fechar exame'...", "INFO")
            btn_fechar_exame = self.wait.until(
                EC.element_to_be_clickable((By.ID, "fecharExameBarraFerramenta"))
            )
            self.driver.execute_script("arguments[0].click();", btn_fechar_exame)
            time.sleep(2)  # Aguardar processamento
            log_message("Botão 'Fechar exame' clicado", "SUCCESS")

            log_message("Máscara preenchida, salva e exame fechado com sucesso!", "SUCCESS")
            return True

        except Exception as e:
            log_message(f"Erro ao preencher máscara: {str(e)}", "ERROR")

            # Garantir que voltamos ao contexto principal em caso de erro
            try:
                self.driver.switch_to.default_content()
                log_message("Voltado para contexto principal após erro", "INFO")
            except:
                pass

            return False

    def run(self, params):
        """Executa a criação de exames para todos os arquivos"""
        username = params.get("username")
        password = params.get("password")
        url = params.get("url", "https://pathoweb.com.br/login/auth")

        # Processar todos os arquivos com OCR
        try:
            log_message("🔍 Iniciando processamento OCR de todos os arquivos...", "INFO")

            arquivos_processados = process_all_files_optimized()

            if not arquivos_processados:
                log_message("❌ Nenhum arquivo foi processado pelo OCR", "ERROR")
                return

            log_message(f"✅ {len(arquivos_processados)} arquivos processados pelo OCR", "SUCCESS")

        except Exception as e:
            log_message(f"❌ Erro no processamento OCR: {str(e)}", "ERROR")
            return

        self.setup(url)

        try:
            log_message("🚀 Iniciando criação de exames Hospital do Câncer...", "INFO")

            # Login
            self.login(username, password)

            # Processar cada arquivo individualmente
            exames_criados = 0
            exames_com_erro = 0

            for i, arquivo_info in enumerate(arquivos_processados, 1):

                log_message(f"📋 Processando exame {i}/{len(arquivos_processados)}: {arquivo_info['arquivo_origem']}", "INFO")

                if self.processar_um_exame(arquivo_info):
                    exames_criados += 1
                else:
                    exames_com_erro += 1

                time.sleep(2)

            # Resumo final
            log_message(f"🎉 Processamento concluído!", "SUCCESS")
            log_message(f"   ✅ Exames criados: {exames_criados}", "SUCCESS")
            if exames_com_erro > 0:
                log_message(f"   ❌ Exames com erro: {exames_com_erro}", "ERROR")

        except Exception as e:
            log_message(f"✗ Erro no processo geral: {str(e)}", "ERROR")
        finally:
            time.sleep(3)
            self.driver.quit()

def run(params):
    module = CriacaoExamesHclSus()
    module.run(params)