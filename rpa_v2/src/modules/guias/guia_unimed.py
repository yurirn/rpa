import os
import time
import pandas as pd
from tkinter import messagebox, filedialog
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from dotenv import load_dotenv
from datetime import datetime

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

load_dotenv()

class GuiaUnimedModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Guia Unimed")

    def get_unique_guias(self, file_path: str) -> list:
        try:
            # Ler a primeira coluna (coluna A) do Excel, onde a primeira linha é o cabeçalho "GUIA"
            df = pd.read_excel(file_path, header=0)
            # Pegar a primeira coluna, ignorando a primeira linha (cabeçalho)
            guias = df.iloc[:, 0].dropna().tolist()
            # Verificar se a primeira linha é cabeçalho (pode ser string "GUIA" ou similar)
            if guias and isinstance(guias[0], str) and guias[0].upper() == "GUIA":
                guias = guias[1:]  # Remove o cabeçalho se for "GUIA"
            # Converter todos os valores para string para garantir compatibilidade
            guias = [str(guia).strip() for guia in guias if str(guia).strip()]
            return guias
        except Exception as e:
            raise ValueError(f"Erro ao ler o Excel: {e}")

    def run(self, params: dict):
        username = params.get("username")
        password = params.get("password")
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode")
        excel_file = params.get("excel_file")

        url = os.getenv("SYSTEM_URL", "https://dap.pathoweb.com.br/login/auth")
        driver = BrowserFactory.create_chrome(headless=headless_mode)
        wait = WebDriverWait(driver, 15)
        # Criar um wait mais longo para operações que podem demorar mais
        wait_long = WebDriverWait(driver, 30)

        try:
            log_message("Iniciando automação de Guia Unimed...", "INFO")

            # Carregar guias do Excel
            if not excel_file or not os.path.exists(excel_file):
                messagebox.showerror("Erro", "Arquivo Excel não informado ou não encontrado.")
                return
            try:
                guias = self.get_unique_guias(excel_file)
            except Exception as e:
                messagebox.showerror("Erro", str(e))
                return
            if not guias:
                messagebox.showerror("Erro", "Nenhuma guia encontrada no arquivo.")
                return
            
            log_message(f"✅ Carregadas {len(guias)} guias do Excel", "SUCCESS")
            
            # Criar DataFrame para armazenar resultados
            resultados_df = pd.DataFrame(columns=["GUIA", "CARTAO", "MEDICO", "CRM", "PROCEDIMENTOS", "QTD", "TEXTO"])

            # Login
            driver.get(url)
            wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
            driver.find_element(By.ID, "password").send_keys(password)
            driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

            log_message("Verificando se precisa navegar para módulo de faturamento...", "INFO")
            current_url = driver.current_url

            if current_url == "https://dap.pathoweb.com.br/" or "trocarModulo" in current_url:
                log_message("Detectada tela de seleção de módulos - navegando para módulo de faturamento...", "INFO")
                try:
                    modulo_link = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=2']")))
                    modulo_link.click()
                    time.sleep(2)
                    log_message("✅ Navegação para módulo de faturamento realizada", "SUCCESS")
                except Exception as e:
                    log_message(f"⚠️ Erro ao navegar para módulo: {e}", "WARNING")
                    driver.get("https://dap.pathoweb.com.br/moduloFaturamento/index")
                    time.sleep(2)
                    log_message("🔄 Navegação direta para módulo realizada", "INFO")

            elif "moduloFaturamento" in current_url:
                log_message("✅ Já está no módulo de faturamento - pulando navegação", "SUCCESS")
            else:
                log_message(f"⚠️ URL inesperada detectada: {current_url}", "WARNING")
                driver.get("https://dap.pathoweb.com.br/moduloFaturamento/index")
                time.sleep(2)
                log_message("🔄 Navegação direta para módulo realizada (fallback)", "INFO")

            try:
                modal_close_button = driver.find_element(By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button")
                if modal_close_button.is_displayed():
                    modal_close_button.click()
                    time.sleep(1)
            except Exception:
                pass

            # Acessar explicitamente a página do módulo de faturamento
            log_message("Acessando módulo de faturamento via URL...", "INFO")
            driver.get("https://dap.pathoweb.com.br/moduloFaturamento/index")

            # Clicar no botão "Preparar exames para fatura"
            log_message("Clicando em 'Preparar exames para fatura'...", "INFO")
            try:
                preparar_btn = wait.until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "a.btn.btn-danger.chamadaAjax.setupAjax[data-url='/moduloFaturamento/preFaturamento']"
                )))
                preparar_btn.click()
            except Exception:
                preparar_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH,
                    "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]"
                )))
                preparar_btn.click()

            # Aguardar possível spinner/modal carregar
            try:
                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                log_message("✅ Modal de carregamento fechado", "INFO")
            except Exception:
                time.sleep(1)

            log_message("Tela de Pré Faturamento aberta.", "SUCCESS")

            # Processar cada guia do Excel
            resultados = []
            for guia in guias:
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break
                try:
                    log_message(f"➡️ Processando guia: {guia}", "INFO")
                    
                    # Digitar o código de barras no campo codigoBarras
                    log_message(f"🔍 Aguardando campo código de barras estar disponível...", "INFO")
                    campo_exame = wait.until(EC.element_to_be_clickable((By.ID, "codigoBarras")))

                    # Aguardar um pouco para garantir que o campo está pronto
                    time.sleep(1)

                    # Limpar e preencher o campo
                    campo_exame.clear()
                    time.sleep(0.5)
                    campo_exame.send_keys(str(guia))
                    log_message(f"✅ Código de barras {guia} digitado no campo", "SUCCESS")
                    time.sleep(0.5)
                    
                    # Clicar no botão Pesquisar
                    pesquisar_btn = wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento")))
                    pesquisar_btn.click()
                    log_message("Pesquisando exame...", "INFO")
                    
                    # Aguardar carregamento dos resultados com mais tempo
                    try:
                        # Primeiro aguardar o spinner aparecer (se existir)
                        try:
                            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                            log_message("🔄 Carregando resultados...", "INFO")
                            wait_long.until(EC.invisibility_of_element_located((By.ID, "spinner")))
                        except Exception:
                            # Se não encontrar o spinner, apenas aguarda um tempo fixo
                            log_message("Aguardando carregamento dos resultados...", "INFO")
                            time.sleep(5)
                    except Exception:
                        log_message("Tempo de carregamento excedido, verificando resultados mesmo assim...", "WARNING")
                    
                    # Aguardar mais um pouco para garantir que a tabela foi carregada
                    time.sleep(3)
                    
                    # Verificar se há resultados usando diferentes seletores
                    tbody_rows = []
                    
                    # Tentar diferentes abordagens para encontrar a tabela de resultados
                    selectors = [
                        "#tabelaPreFaturamentoTbody tr",
                        ".table-responsive table tbody tr",
                        "table.table-striped tbody tr",
                        "table.footable tbody tr"
                    ]
                    
                    for selector in selectors:
                        try:
                            tbody_rows = driver.find_elements(By.CSS_SELECTOR, selector)
                            if len(tbody_rows) > 0:
                                log_message(f"Tabela de resultados encontrada usando seletor: {selector}", "INFO")
                                break
                        except Exception:
                            continue
                    
                    # Se ainda não encontrou resultados, tenta verificar se há mensagem de "nenhum resultado"
                    if len(tbody_rows) == 0:
                        try:
                            # Verificar se há mensagem de "nenhum resultado"
                            no_results_msg = driver.find_element(By.XPATH, "//*[contains(text(), 'Nenhum resultado encontrado')]")
                            if no_results_msg:
                                log_message(f"⚠️ Mensagem de 'Nenhum resultado encontrado' para {guia}", "WARNING")
                        except Exception:
                            # Se não encontrar a mensagem, aguarda mais um pouco e tenta novamente
                            log_message("Aguardando mais tempo para carregamento completo...", "INFO")
                            time.sleep(5)
                            for selector in selectors:
                                try:
                                    tbody_rows = driver.find_elements(By.CSS_SELECTOR, selector)
                                    if len(tbody_rows) > 0:
                                        log_message(f"Tabela de resultados encontrada após espera adicional", "INFO")
                                        break
                                except Exception:
                                    continue
                    
                    if len(tbody_rows) == 0:
                        log_message(f"⚠️ Nenhum resultado encontrado para {guia}. Pulando.", "WARNING")
                        resultados.append({"guia": guia, "status": "sem_resultados"})
                        # Adicionar linha vazia no DataFrame
                        resultados_df = pd.concat([resultados_df, pd.DataFrame([{
                            "GUIA": guia,
                            "CARTAO": "",
                            "MEDICO": "",
                            "CRM": "",
                            "PROCEDIMENTOS": "",
                            "QTD": "",
                            "TEXTO": ""
                        }])], ignore_index=True)
                        continue
                    
                    log_message(f"✅ Encontrados {len(tbody_rows)} resultados para a guia {guia}", "SUCCESS")
                    
                    # Processar primeira linha para obter dados básicos
                    try:
                        # Inicializar variáveis
                        cartao = ""
                        medico = ""
                        crm = ""
                        texto = ""
                        procedimentos_str = ""
                        quantidades_str = ""
                        
                        # Obter número do cartão do paciente da tabela
                        try:
                            # Índice pode variar dependendo da estrutura da tabela
                            # Tentar diferentes índices para o cartão
                            try:
                                cartao = tbody_rows[0].find_elements(By.CSS_SELECTOR, "td")[6].text.strip()
                            except:
                                try:
                                    # Tentar outro índice comum para o campo de cartão
                                    cartao = tbody_rows[0].find_elements(By.CSS_SELECTOR, "td")[5].text.strip()
                                except:
                                    # Se ainda falhar, tentar localizar pela coluna "Carteira"
                                    header_cells = driver.find_elements(By.CSS_SELECTOR, "table th")
                                    cartao_index = -1
                                    for i, cell in enumerate(header_cells):
                                        if "carteira" in cell.text.lower():
                                            cartao_index = i
                                            break
                                    
                                    if cartao_index >= 0:
                                        cartao = tbody_rows[0].find_elements(By.CSS_SELECTOR, "td")[cartao_index].text.strip()
                            
                            log_message(f"✅ Número do cartão obtido: {cartao}", "INFO")
                        except Exception as e:
                            log_message(f"⚠️ Erro ao obter número do cartão: {e}", "WARNING")
                        
                        # NOVO FLUXO: Marcar checkbox do primeiro exame e clicar no botão "Abrir exame"
                        log_message("Marcando checkbox do primeiro exame...", "INFO")
                        
                        # Encontrar e marcar o checkbox do primeiro exame
                        try:
                            checkbox = tbody_rows[0].find_element(By.CSS_SELECTOR, "input[type='checkbox'][name='exameId']")
                            if not checkbox.is_selected():
                                checkbox.click()
                                log_message("✅ Checkbox do exame marcado", "SUCCESS")
                            else:
                                log_message("ℹ️ Checkbox já estava marcado", "INFO")
                            
                            # Aguardar um pouco após marcar o checkbox
                            time.sleep(1)
                            
                            # Procurar e clicar no botão "Abrir exame"
                            log_message("Procurando botão 'Abrir exame'...", "INFO")
                            
                            try:
                                # Procurar pelo botão "Abrir exame" usando o seletor específico
                                abrir_btn = wait.until(EC.element_to_be_clickable((
                                    By.CSS_SELECTOR, 
                                    "a.btn.btn-sm.btn-primary.chamadaAjax.toogleInicial.setupAjax[data-url='/moduloFaturamento/abrirExameCorrecao']"
                                )))
                                log_message("✅ Botão 'Abrir exame' encontrado", "SUCCESS")
                                
                                # Clicar no botão
                                abrir_btn.click()
                                log_message("✅ Clique no botão 'Abrir exame' realizado", "SUCCESS")
                                
                                # Aguardar o modal aparecer
                                log_message("Aguardando modal do exame abrir...", "INFO")
                                time.sleep(3)
                                
                                # Verificar se o modal foi aberto
                                try:
                                    modal = wait.until(EC.presence_of_element_located((By.ID, "myModal")))
                                    if modal.is_displayed():
                                        log_message("✅ Modal do exame aberto com sucesso", "SUCCESS")
                                    else:
                                        log_message("⚠️ Modal encontrado mas não está visível", "WARNING")
                                        time.sleep(2)  # Aguardar mais um pouco
                                except Exception:
                                    log_message("⚠️ Modal não encontrado, tentando continuar...", "WARNING")
                                    time.sleep(2)
                                
                            except Exception as e:
                                log_message(f"❌ Erro ao clicar no botão 'Abrir exame': {e}", "ERROR")
                                raise Exception(f"Não foi possível abrir o exame: {e}")
                                
                        except Exception as e:
                            log_message(f"❌ Erro ao marcar checkbox do exame: {e}", "ERROR")
                            raise Exception(f"Não foi possível marcar o exame: {e}")
                        
                        # Extrair nome do médico e CRM do modal aberto
                        try:
                            # Método 1: Usar JavaScript para extrair o valor do input (mais confiável)
                            try:
                                medico = driver.execute_script("return $('#medicoRequisitanteInput').val();")
                                if medico and medico.strip():
                                    medico = medico.strip()
                                    log_message(f"✅ Médico requisitante encontrado (JavaScript): {medico}", "SUCCESS")
                                else:
                                    raise Exception("Valor vazio retornado pelo JavaScript")
                            except Exception as e:
                                log_message(f"Tentando método alternativo para médico: {e}", "INFO")
                                
                                # Método 2: Procurar diretamente pelo input, mesmo que esteja oculto
                                try:
                                    medico_input = driver.find_element(By.ID, "medicoRequisitanteInput")
                                    medico = medico_input.get_attribute("value").strip()
                                    if medico:
                                        log_message(f"✅ Médico requisitante encontrado (input direto): {medico}", "SUCCESS")
                                    else:
                                        raise Exception("Input encontrado mas valor vazio")
                                except Exception:
                                    # Método 3: Procurar pelo elemento <a> com a classe "table-editable-ancora"
                                    try:
                                        medico_element = driver.find_element(By.CSS_SELECTOR, 
                                            "a.table-editable-ancora.autocomplete.autocompleteSetup")
                                        medico = medico_element.text.strip()
                                        if medico:
                                            log_message(f"✅ Médico requisitante encontrado (link ancora): {medico}", "SUCCESS")
                                        else:
                                            raise Exception("Link encontrado mas texto vazio")
                                    except Exception:
                                        # Método 4: Procurar qualquer elemento após "Médico requisitante"
                                        try:
                                            # Localizar o elemento td que contém "Médico requisitante"
                                            medico_label = driver.find_element(By.XPATH, "//td[contains(text(), 'Médico requisitante')]")
                                            # Pegar o elemento irmão (following-sibling)
                                            medico_td = medico_label.find_element(By.XPATH, "following-sibling::td")
                                            # Extrair o texto completo do elemento
                                            medico = medico_td.text.strip()
                                            if medico:
                                                log_message(f"✅ Médico requisitante encontrado (texto do td): {medico}", "SUCCESS")
                                            else:
                                                raise Exception("TD encontrado mas texto vazio")
                                        except Exception:
                                            # Método 5: Usar JavaScript alternativo para procurar o elemento
                                            try:
                                                medico = driver.execute_script("""
                                                    var input = document.getElementById('medicoRequisitanteInput');
                                                    if (input && input.value) {
                                                        return input.value;
                                                    }
                                                    var ancora = document.querySelector('a.table-editable-ancora.autocomplete.autocompleteSetup');
                                                    if (ancora && ancora.textContent) {
                                                        return ancora.textContent.trim();
                                                    }
                                                    return null;
                                                """)
                                                if medico and medico.strip():
                                                    medico = medico.strip()
                                                    log_message(f"✅ Médico requisitante encontrado (JavaScript alternativo): {medico}", "SUCCESS")
                                                else:
                                                    raise Exception("JavaScript alternativo não retornou resultado")
                                            except Exception:
                                                log_message("⚠️ Todos os métodos falharam para encontrar o médico", "WARNING")
                            
                            # Extrair CRM do typeahead dropdown
                            try:
                                log_message("Extraindo CRM do médico...", "INFO")
                                
                                # Função helper para verificar se dropdown está pronto
                                def dropdown_pronto():
                                    try:
                                        dropdown = driver.find_element(By.CSS_SELECTOR, "ul.typeahead li.active a")
                                        return dropdown.is_displayed() and "CRM:" in dropdown.text
                                    except:
                                        return False
                                
                                # Função helper para aguardar condição com polling rápido
                                def aguardar_condicao(condicao_func, timeout=5, intervalo=0.1):
                                    import time
                                    start_time = time.time()
                                    while time.time() - start_time < timeout:
                                        if condicao_func():
                                            return True
                                        time.sleep(intervalo)
                                    return False
                                
                                # Aguardar tabela aparecer com polling rápido
                                def tabela_pronta():
                                    try:
                                        return driver.find_element(By.ID, "requisicao_r").is_displayed()
                                    except:
                                        return False
                                
                                if not aguardar_condicao(tabela_pronta, timeout=8):
                                    raise Exception("Tabela não carregou")
                                
                                # Verificar se dropdown já está visível
                                if dropdown_pronto():
                                    log_message("✅ Dropdown já visível!", "SUCCESS")
                                else:
                                    # Tentar ativar dropdown
                                    ativado = False
                                    
                                    # Método 1: Input
                                    try:
                                        def input_pronto():
                                            try:
                                                input_elem = driver.find_element(By.CSS_SELECTOR, "#requisicao_r #medicoRequisitanteInput")
                                                return input_elem.is_displayed() and input_elem.is_enabled()
                                            except:
                                                return False
                                        
                                        if aguardar_condicao(input_pronto, timeout=3):
                                            medico_input = driver.find_element(By.CSS_SELECTOR, "#requisicao_r #medicoRequisitanteInput")
                                            medico_input.click()
                                            
                                            if aguardar_condicao(dropdown_pronto, timeout=2):
                                                log_message("✅ Dropdown ativado via input", "SUCCESS")
                                                ativado = True
                                    except:
                                        pass
                                    
                                    # Método 2: Âncora (se input falhou)
                                    if not ativado:
                                        try:
                                            def ancora_pronta():
                                                try:
                                                    ancora = driver.find_element(By.CSS_SELECTOR, "#requisicao_r a.table-editable-ancora.autocomplete.autocompleteSetup")
                                                    return ancora.is_displayed() and ancora.is_enabled()
                                                except:
                                                    return False
                                            
                                            if aguardar_condicao(ancora_pronta, timeout=2):
                                                ancora = driver.find_element(By.CSS_SELECTOR, "#requisicao_r a.table-editable-ancora.autocomplete.autocompleteSetup")
                                                ancora.click()
                                                
                                                if aguardar_condicao(dropdown_pronto, timeout=2):
                                                    log_message("✅ Dropdown ativado via âncora", "SUCCESS")
                                                    ativado = True
                                        except:
                                            pass
                                    
                                    if not ativado:
                                        log_message("⚠️ Não conseguiu ativar dropdown", "WARNING")
                                
                                # Extrair CRM do dropdown (método otimizado)
                                crm = ""
                                try:
                                    # Método JavaScript mais rápido
                                    crm = driver.execute_script("""
                                        try {
                                            let crmElement = document.querySelector("ul.typeahead li.active a");
                                            if (crmElement && crmElement.innerText) {
                                                let crmText = crmElement.innerText;
                                                let crmMatch = crmText.match(/CRM:\\s*(\\S+)/);
                                                return crmMatch ? crmMatch[1] : null;
                                            }
                                        } catch (e) {}
                                        return null;
                                    """)
                                    
                                    if crm:
                                        log_message(f"✅ CRM encontrado: {crm}", "SUCCESS")
                                    else:
                                        # Fallback direto sem delay
                                        try:
                                            dropdown_elem = driver.find_element(By.CSS_SELECTOR, "ul.typeahead li.active a")
                                            crm_text = dropdown_elem.text
                                            import re
                                            crm_match = re.search(r'CRM:\s*(\S+)', crm_text)
                                            if crm_match:
                                                crm = crm_match.group(1)
                                                log_message(f"✅ CRM extraído: {crm}", "SUCCESS")
                                        except:
                                            log_message("⚠️ CRM não encontrado", "WARNING")
                                
                                except Exception as e:
                                    log_message(f"⚠️ Erro ao extrair CRM: {e}", "WARNING")
                                
                                # Fechar dropdown rapidamente
                                try:
                                    driver.execute_script("document.body.click();")
                                except:
                                    pass
                                    
                            except Exception as e:
                                log_message(f"⚠️ Erro ao extrair CRM: {e}", "WARNING")
                            
                            if not medico:
                                log_message("⚠️ Não foi possível encontrar o médico requisitante", "WARNING")
                            if not crm:
                                log_message("⚠️ Não foi possível encontrar o CRM", "WARNING")
                        except Exception as e:
                            log_message(f"⚠️ Erro ao obter médico requisitante: {e}", "WARNING")
                        
                        # Extrair procedimentos e quantidades do modal
                        procedimentos = []
                        quantidades = []
                        try:
                            log_message("Extraindo procedimentos do modal...", "INFO")
                            
                            # Aguardar a div de procedimentos estar presente no modal
                            wait.until(EC.presence_of_element_located((By.ID, "divProcedimentos")))
                            
                            # Encontrar todas as linhas de procedimentos na tabela
                            # Seleciona todas as linhas tr que têm id começando com "procedimento_" mas não "novosProcedimentos"
                            procedimento_rows = driver.find_elements(By.CSS_SELECTOR, "#divProcedimentos table tbody tr[id^='procedimento_']:not(#novosProcedimentos)")
                            
                            if not procedimento_rows:
                                # Tentar método alternativo sem o filtro :not
                                procedimento_rows = driver.find_elements(By.CSS_SELECTOR, "#divProcedimentos table tbody tr[id^='procedimento_']")
                                # Remover a linha "novosProcedimentos" se estiver presente
                                procedimento_rows = [row for row in procedimento_rows if row.get_attribute("id") != "novosProcedimentos"]
                            
                            if not procedimento_rows:
                                log_message("⚠️ Nenhuma linha de procedimento encontrada, tentando método alternativo...", "WARNING")
                                # Método alternativo: buscar todas as linhas da tabela exceto cabeçalho
                                procedimento_rows = driver.find_elements(By.CSS_SELECTOR, "#divProcedimentos table tbody tr")
                                # Filtrar apenas as que têm checkbox de procedimento
                                procedimento_rows = [row for row in procedimento_rows if row.find_elements(By.CSS_SELECTOR, "input[type='checkbox'][name='procedimentoExameId']")]
                            
                            log_message(f"✅ Encontradas {len(procedimento_rows)} linhas de procedimentos", "SUCCESS")
                            
                            for row in procedimento_rows:
                                try:
                                    # Extrair código do procedimento (apenas a parte antes do " -")
                                    # O nome está em um link <a> com classe "table-editable-ancora autocomplete autocompleteSetup" na coluna "Nome"
                                    procedimento_codigo = ""
                                    try:
                                        # Tentar encontrar o link com o nome do procedimento
                                        procedimento_link = row.find_element(By.CSS_SELECTOR, "td:nth-child(3) a.table-editable-ancora.autocomplete.autocompleteSetup")
                                        procedimento_texto = procedimento_link.text.strip()
                                        
                                        # Se não encontrar, tentar alternativa
                                        if not procedimento_texto or procedimento_texto == "Vazio":
                                            # Tentar pelo input oculto
                                            procedimento_input = row.find_element(By.CSS_SELECTOR, "td:nth-child(3) input.autocomplete")
                                            procedimento_texto = procedimento_input.get_attribute("value").strip()
                                        
                                        # Extrair apenas o código (parte antes do " -")
                                        if procedimento_texto and " -" in procedimento_texto:
                                            procedimento_codigo = procedimento_texto.split(" -")[0].strip()
                                        elif procedimento_texto:
                                            # Se não tiver " -", usar o texto inteiro (caso seja só o código)
                                            procedimento_codigo = procedimento_texto.strip()
                                        
                                    except Exception as e:
                                        log_message(f"⚠️ Erro ao extrair código do procedimento: {e}", "WARNING")
                                        # Tentar método alternativo: pegar texto direto da célula
                                        try:
                                            cells = row.find_elements(By.CSS_SELECTOR, "td")
                                            if len(cells) >= 3:
                                                procedimento_texto = cells[2].text.strip()
                                                # Extrair apenas o código (parte antes do " -")
                                                if procedimento_texto and " -" in procedimento_texto:
                                                    procedimento_codigo = procedimento_texto.split(" -")[0].strip()
                                                elif procedimento_texto:
                                                    procedimento_codigo = procedimento_texto.strip()
                                        except:
                                            procedimento_codigo = ""
                                    
                                    # Extrair quantidade
                                    try:
                                        # A quantidade está na segunda coluna (índice 1)
                                        quantidade_link = row.find_element(By.CSS_SELECTOR, "td:nth-child(2) a.table-editable-ancora")
                                        quantidade = quantidade_link.text.strip()
                                        
                                        # Se não encontrar, tentar pelo input
                                        if not quantidade or quantidade == "":
                                            quantidade_input = row.find_element(By.CSS_SELECTOR, "td:nth-child(2) input[type='number']")
                                            quantidade = quantidade_input.get_attribute("value").strip()
                                        
                                    except Exception as e:
                                        log_message(f"⚠️ Erro ao extrair quantidade: {e}", "WARNING")
                                        # Tentar método alternativo
                                        try:
                                            cells = row.find_elements(By.CSS_SELECTOR, "td")
                                            if len(cells) >= 2:
                                                quantidade = cells[1].text.strip()
                                            else:
                                                quantidade = "1"
                                        except:
                                            quantidade = "1"
                                    
                                    # Só adicionar se o código do procedimento não for vazio
                                    if procedimento_codigo and procedimento_codigo != "Vazio" and procedimento_codigo != "":
                                        procedimentos.append(procedimento_codigo)
                                        quantidades.append(quantidade if quantidade else "1")
                                        log_message(f"✅ Procedimento encontrado: {procedimento_codigo} - Qtd: {quantidade}", "INFO")
                                    
                                except Exception as e:
                                    log_message(f"⚠️ Erro ao processar linha de procedimento: {e}", "WARNING")
                                    continue
                            
                            # Formatar strings finais
                            procedimentos_str = ", ".join(procedimentos) if procedimentos else ""
                            quantidades_str = ", ".join(quantidades) if quantidades else ""
                            
                            if procedimentos_str:
                                log_message(f"✅ Procedimentos obtidos: {procedimentos_str}", "SUCCESS")
                                log_message(f"✅ Quantidades obtidas: {quantidades_str}", "SUCCESS")
                            else:
                                log_message("⚠️ Nenhum procedimento válido encontrado", "WARNING")
                                
                        except Exception as e:
                            log_message(f"⚠️ Erro ao extrair procedimentos do modal: {e}", "WARNING")
                            procedimentos_str = ""
                            quantidades_str = ""
                        
                        # Extrair texto clínico do modal
                        texto = ""
                        try:
                            # Primeiro tentar localizar o iframe dentro do modal
                            try:
                                iframe = driver.find_element(By.CSS_SELECTOR, "#myModal .cke_wysiwyg_frame")
                                driver.switch_to.frame(iframe)
                                
                                # Agora obter o texto do corpo do iframe
                                texto_element = driver.find_element(By.CSS_SELECTOR, "body")
                                texto = texto_element.text.strip()
                                
                                # Voltar ao contexto principal
                                driver.switch_to.default_content()
                            except:
                                # Se não encontrar o iframe, tentar outros seletores para o texto clínico dentro do modal
                                try:
                                    texto_element = driver.find_element(
                                        By.XPATH, 
                                        "//div[@id='myModal']//*[contains(text(), 'Dados clínicos')]/following-sibling::*"
                                    )
                                    texto = texto_element.text.strip()
                                except:
                                    # Última tentativa - procurar por div ou textarea com conteúdo dentro do modal
                                    elements = driver.find_elements(By.CSS_SELECTOR, "#myModal div.form-control, #myModal textarea.form-control")
                                    for elem in elements:
                                        if elem.text and len(elem.text) > 5:
                                            texto = elem.text.strip()
                                            break
                            
                            log_message(f"✅ Texto clínico obtido: {texto[:50]}...", "INFO")
                        except Exception as e:
                            log_message(f"⚠️ Erro ao obter texto clínico: {e}", "WARNING")
                        
                        # Fechar modal
                        try:
                            # Procurar botão de fechar modal
                            close_btn = driver.find_element(By.CSS_SELECTOR, "#myModal .modal-header .close")
                            close_btn.click()
                            time.sleep(1)
                            log_message("✅ Modal fechado", "INFO")
                        except:
                            # Tentar fechar com ESC
                            try:
                                from selenium.webdriver.common.keys import Keys
                                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                                time.sleep(1)
                                log_message("✅ Modal fechado com ESC", "INFO")
                            except:
                                log_message("⚠️ Não foi possível fechar o modal", "WARNING")
                        
                        # Adicionar dados ao DataFrame
                        resultados_df = pd.concat([resultados_df, pd.DataFrame([{
                            "GUIA": guia,
                            "CARTAO": cartao,
                            "MEDICO": medico,
                            "CRM": crm,
                            "PROCEDIMENTOS": procedimentos_str,
                            "QTD": quantidades_str,
                            "TEXTO": texto
                        }])], ignore_index=True)
                        
                        resultados.append({"guia": guia, "status": "sucesso"})
                        log_message(f"✅ Guia {guia} processada com sucesso", "SUCCESS")
                        
                    except Exception as e:
                        log_message(f"❌ Erro ao processar detalhes da guia {guia}: {e}", "ERROR")
                        resultados.append({"guia": guia, "status": "erro_detalhes", "erro": str(e)})
                        
                        # Adicionar linha com dados parciais no DataFrame
                        resultados_df = pd.concat([resultados_df, pd.DataFrame([{
                            "GUIA": guia,
                            "CARTAO": cartao if 'cartao' in locals() else "",
                            "MEDICO": "",
                            "CRM": "",
                            "PROCEDIMENTOS": procedimentos_str if 'procedimentos_str' in locals() else "",
                            "QTD": quantidades_str if 'quantidades_str' in locals() else "",
                            "TEXTO": ""
                        }])], ignore_index=True)
                    
                except Exception as e:
                    resultados.append({"guia": guia, "status": "erro", "erro": str(e)})
                    log_message(f"❌ Erro ao processar guia {guia}: {e}", "ERROR")
                    
                    # Adicionar linha vazia no DataFrame para a guia com erro
                    resultados_df = pd.concat([resultados_df, pd.DataFrame([{
                        "GUIA": guia,
                        "CARTAO": "",
                        "MEDICO": "",
                        "CRM": "",
                        "PROCEDIMENTOS": "",
                        "QTD": "",
                        "TEXTO": ""
                    }])], ignore_index=True)

            # Salvar resultados em Excel
            try:
                # Gerar nome do arquivo com timestamp
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_dir = os.path.dirname(excel_file)
                output_file = os.path.join(output_dir, f"resultados_guias_unimed_{timestamp}.xlsx")
                
                # Salvar DataFrame para Excel
                resultados_df.to_excel(output_file, index=False)
                log_message(f"✅ Resultados salvos em: {output_file}", "SUCCESS")
            except Exception as e:
                log_message(f"❌ Erro ao salvar arquivo de resultados: {e}", "ERROR")

            # Resumo final
            total = len(resultados)
            sucesso = [r for r in resultados if r["status"] == "sucesso"]
            erro = [r for r in resultados if r["status"] in ["erro", "erro_detalhes", "erro_link"]]
            sem_resultados = [r for r in resultados if r["status"] == "sem_resultados"]
            
            log_message("\nResumo do processamento:", "INFO")
            log_message(f"Total de guias: {total}", "INFO")
            log_message(f"Processadas com sucesso: {len(sucesso)}", "SUCCESS")
            log_message(f"Sem resultados: {len(sem_resultados)}", "WARNING")
            log_message(f"Erros: {len(erro)}", "ERROR")
            
            messagebox.showinfo("Sucesso",
                f"✅ Processamento finalizado!\n"
                f"Total: {total}\n"
                f"Sucesso: {len(sucesso)}\n"
                f"Sem resultados: {len(sem_resultados)}\n"
                f"Erros: {len(erro)}\n\n"
                f"Resultados salvos em:\n{output_file if 'output_file' in locals() else 'Erro ao salvar arquivo'}"
            )

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            driver.quit()


def run(params: dict):
    module = GuiaUnimedModule()
    module.run(params)
