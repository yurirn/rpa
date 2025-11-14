import os
import time
import pandas as pd
import requests
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
from datetime import datetime
import re

from urllib.parse import parse_qs, unquote, urljoin, urlparse
from requests.utils import add_dict_to_cookiejar

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

load_dotenv()

class FaturaMensalModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Fatura Mensal")

    def get_excel_data(self, file_path: str) -> list:
        """Lê dados do Excel e retorna lista de dicionários com CLIENTE, TIPO e DATA"""
        try:
            df = pd.read_excel(file_path, header=0)
            
            # Verificar se tem as colunas necessárias
            if len(df.columns) < 3:
                raise ValueError("Excel deve ter pelo menos 3 colunas: CLIENTE, TIPO, DATA")
            
            # Pegar as primeiras 3 colunas
            clientes = df.iloc[:, 0].dropna().tolist()
            tipos = df.iloc[:, 1].dropna().tolist()
            datas = df.iloc[:, 2].dropna().tolist()
            
            # Verificar se a primeira linha é cabeçalho
            if clientes and isinstance(clientes[0], str) and clientes[0].upper() in ["CLIENTE", "CONVENIO", "PROCEDENCIA"]:
                clientes = clientes[1:]
                tipos = tipos[1:]
                datas = datas[1:]
            
            # Converter para string e limpar
            clientes = [str(c).strip() for c in clientes if str(c).strip()]
            tipos = [str(t).strip() for t in tipos if str(t).strip()]
            datas = [str(d).strip() for d in datas if str(d).strip()]
            
            # Criar lista de dicionários
            dados = []
            for i in range(len(clientes)):
                dados.append({
                    'cliente': clientes[i],
                    'tipo': tipos[i] if i < len(tipos) else '',
                    'data': datas[i] if i < len(datas) else ''
                })
            
            return dados
        except Exception as e:
            raise ValueError(f"Erro ao ler o Excel: {e}")

    def parse_date_range(self, date_str: str) -> tuple:
        """Converte string de data no formato '20/09 – 20/10' para tupla (data_inicio, data_fim) com ano completo"""
        try:
            # Remover espaços e dividir por '–' ou '-'
            date_str = date_str.strip()
            if '–' in date_str:
                parts = date_str.split('–')
            elif '-' in date_str:
                parts = date_str.split('-')
            else:
                raise ValueError("Formato de data inválido")
            
            if len(parts) != 2:
                raise ValueError("Formato de data inválido")
            
            data_inicio = parts[0].strip()
            data_fim = parts[1].strip()
            
            # Validar formato DD/MM
            pattern = r'^\d{2}/\d{2}$'
            if not re.match(pattern, data_inicio) or not re.match(pattern, data_fim):
                raise ValueError("Formato de data deve ser DD/MM")
            
            # Adicionar ano atual
            from datetime import datetime
            ano_atual = datetime.now().year
            
            # Converter para formato DD/MM/YYYY
            data_inicio_completa = f"{data_inicio}/{ano_atual}"
            data_fim_completa = f"{data_fim}/{ano_atual}"
            
            return data_inicio_completa, data_fim_completa
        except Exception as e:
            raise ValueError(f"Erro ao processar data '{date_str}': {e}")

    def find_option_by_text(self, driver, select_element, text_to_find: str) -> str:
        """Encontra o valor da opção baseado no texto"""
        try:
            options = select_element.find_elements(By.TAG_NAME, "option")
            log_message(f"🔍 Procurando '{text_to_find}' entre {len(options)} opções", "INFO")
            
            # Busca exata primeiro
            for option in options:
                option_text = option.text.strip()
                if text_to_find.upper() == option_text.upper():
                    log_message(f"✅ Encontrado match exato: '{option_text}' = '{text_to_find}'", "SUCCESS")
                    return option.get_attribute("value")
            
            # Busca parcial (contém)
            for option in options:
                option_text = option.text.strip()
                if text_to_find.upper() in option_text.upper():
                    log_message(f"✅ Encontrado match parcial: '{option_text}' contém '{text_to_find}'", "SUCCESS")
                    return option.get_attribute("value")
            
            # Busca reversa (está contido)
            for option in options:
                option_text = option.text.strip()
                if option_text.upper() in text_to_find.upper():
                    log_message(f"✅ Encontrado match reverso: '{text_to_find}' contém '{option_text}'", "SUCCESS")
                    return option.get_attribute("value")
            
            # Busca por palavras individuais
            palavras_busca = text_to_find.upper().split()
            for option in options:
                option_text = option.text.strip()
                palavras_opcao = option_text.upper().split()
                if any(palavra in palavras_opcao for palavra in palavras_busca):
                    log_message(f"✅ Encontrado match por palavra: '{option_text}'", "SUCCESS")
                    return option.get_attribute("value")
            
            log_message(f"⚠️ Nenhum match encontrado para '{text_to_find}'", "WARNING")
            return ""
        except Exception as e:
            log_message(f"Erro ao buscar opção '{text_to_find}': {e}", "WARNING")
            return ""

    def select_select2_option(self, driver, field_id: str, text_to_find: str) -> bool:
        """Seleciona opção em campo Select2 - digita no campo de busca para filtrar"""
        try:
            log_message(f"🔍 Abrindo Select2 '{field_id}' e procurando '{text_to_find}'", "INFO")
            
            # Primeiro, tentar abrir o Select2 usando JavaScript (mais confiável)
            try:
                driver.execute_script(f"$('#{field_id}').select2('open');")
                log_message(f"✅ Campo Select2 '{field_id}' aberto via JavaScript select2('open')", "SUCCESS")
                time.sleep(2)  # Aguardar dropdown abrir completamente
            except Exception as e1:
                log_message(f"⚠️ JavaScript select2('open') falhou: {e1}", "WARNING")
                # Se falhar, tentar clicar no span do Select2
                try:
                    select2_selection = driver.find_element(By.CSS_SELECTOR, f"#select2-{field_id}-container")
                    driver.execute_script("arguments[0].click();", select2_selection)
                    log_message(f"✅ Campo Select2 '{field_id}' clicado via JavaScript click", "SUCCESS")
                    time.sleep(2)
                except Exception as e2:
                    log_message(f"⚠️ JavaScript click falhou: {e2}", "WARNING")
                    # Última tentativa: clique normal
                    try:
                        select2_selection = driver.find_element(By.CSS_SELECTOR, f"#select2-{field_id}-container")
                        select2_selection.click()
                        log_message(f"✅ Campo Select2 '{field_id}' clicado normalmente", "SUCCESS")
                        time.sleep(2)
                    except Exception as e3:
                        log_message(f"⚠️ Erro ao abrir Select2 '{field_id}': {e3}", "WARNING")
                        return False
            
            # Agora digitar no campo de busca do Select2
            try:
                search_field = driver.find_element(By.CSS_SELECTOR, f"#select2-{field_id}-results + .select2-search .select2-search__field")
                search_field.clear()
                search_field.send_keys(text_to_find)
                log_message(f"✅ Digitado '{text_to_find}' no campo de busca", "SUCCESS")
                time.sleep(1)  # Aguardar filtro processar
            except Exception:
                # Tentar seletor alternativo
                try:
                    search_field = driver.find_element(By.CSS_SELECTOR, ".select2-search__field")
                    search_field.clear()
                    search_field.send_keys(text_to_find)
                    log_message(f"✅ Digitado '{text_to_find}' no campo de busca (seletor alternativo)", "SUCCESS")
                    time.sleep(1)
                except Exception:
                    # Usar JavaScript para digitar
                    try:
                        driver.execute_script(f"$('.select2-search__field').val('{text_to_find}').trigger('input');")
                        log_message(f"✅ Digitado '{text_to_find}' no campo de busca via JavaScript", "SUCCESS")
                        time.sleep(1)
                    except Exception as e:
                        log_message(f"⚠️ Erro ao digitar no campo de busca: {e}", "WARNING")
                        return False
            
            # Aguardar um pouco para o filtro processar
            time.sleep(1)
            
            # Procurar nas opções filtradas do Select2
            select2_options = driver.find_elements(By.CSS_SELECTOR, f"#select2-{field_id}-results .select2-results__option")
            log_message(f"🔍 Encontradas {len(select2_options)} opções filtradas no Select2 '{field_id}'", "INFO")
            
            # Se encontrou opções, clicar na primeira (que deve ser a correta após o filtro)
            if select2_options:
                first_option = select2_options[0]
                option_text = first_option.text.strip()
                log_message(f"✅ Clicando na primeira opção filtrada: '{option_text}'", "SUCCESS")
                first_option.click()
                time.sleep(1)
                return True
            
            # Se não encontrou opções filtradas, tentar busca manual
            for option in select2_options:
                option_text = option.text.strip()
                if text_to_find.upper() == option_text.upper():
                    log_message(f"✅ Encontrado match exato no Select2 '{field_id}': '{option_text}'", "SUCCESS")
                    option.click()
                    time.sleep(1)
                    return True
            
            # Busca parcial (contém)
            for option in select2_options:
                option_text = option.text.strip()
                if text_to_find.upper() in option_text.upper():
                    log_message(f"✅ Encontrado match parcial no Select2 '{field_id}': '{option_text}'", "SUCCESS")
                    option.click()
                    time.sleep(1)
                    return True
            
            log_message(f"⚠️ Nenhum match encontrado no Select2 '{field_id}' para '{text_to_find}'", "WARNING")
            return False
            
        except Exception as e:
            log_message(f"Erro ao selecionar opção Select2 '{text_to_find}' no campo '{field_id}': {e}", "WARNING")
            return False

    def wait_for_download(self, download_dir: str, previous_files: set, timeout: int = 45) -> str:
        """Aguarda a criação de um novo arquivo PDF no diretório de downloads"""
        try:
            if not os.path.isdir(download_dir):
                return ""

            end_time = time.time() + timeout
            while time.time() < end_time:
                current_files = set(os.listdir(download_dir))
                new_files = [
                    f for f in current_files - previous_files
                    if f.lower().endswith(".pdf")
                ]

                if new_files:
                    newest_file = max(
                        new_files,
                        key=lambda f: os.path.getmtime(os.path.join(download_dir, f))
                    )
                    file_path = os.path.join(download_dir, newest_file)
                    if os.path.exists(file_path):
                        return file_path

                time.sleep(1)
        except Exception as e:
            log_message(f"⚠️ Erro ao monitorar downloads: {e}", "WARNING")

        return ""

    def get_pdf_url(self, driver, base_url: str) -> str:
        """Obtém a URL absoluta do PDF a partir da aba atual"""
        try:
            current_url = driver.current_url or ""
            if current_url and current_url != "about:blank":
                if "renderReport" in current_url:
                    return current_url

            log_message("⚠️ URL da aba PDF é 'about:blank'. Tentando recuperar link do PDF no DOM...", "WARNING")
            wait = WebDriverWait(driver, 15)

            # Aguardar página carregar completamente
            time.sleep(2)
            
            # Tentar localizar o link <a> que contém o botão "Abrir" - método mais confiável
            try:
                # Primeiro tentar encontrar o link que contém o botão com ID "open-button"
                link_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[href*='renderReport']")))
                href = link_element.get_attribute("href")
                if href:
                    pdf_url = urljoin(base_url, href) if not href.startswith("http") else href
                    log_message(f"✅ URL obtida via link <a>: {pdf_url}", "SUCCESS")
                    return pdf_url
            except Exception as e_link1:
                log_message(f"ℹ️ Link <a> com renderReport não encontrado (tentativa 1): {e_link1}", "INFO")
            
            # Tentar encontrar qualquer link que contenha renderReport
            try:
                links = driver.find_elements(By.TAG_NAME, "a")
                for link in links:
                    href = link.get_attribute("href")
                    if href and "renderReport" in href:
                        pdf_url = urljoin(base_url, href) if not href.startswith("http") else href
                        log_message(f"✅ URL obtida via busca em links: {pdf_url}", "SUCCESS")
                        return pdf_url
            except Exception as e_link2:
                log_message(f"ℹ️ Busca em links falhou: {e_link2}", "INFO")

            # Tentar localizar elemento <embed>
            try:
                embed_element = wait.until(EC.presence_of_element_located((By.TAG_NAME, "embed")))
                embed_src = embed_element.get_attribute("src") or embed_element.get_attribute("data")

                if embed_src:
                    pdf_url = urljoin(base_url, embed_src)
                    log_message(f"✅ URL obtida via <embed>: {pdf_url}", "SUCCESS")
                    return pdf_url
            except Exception as e_embed:
                log_message(f"ℹ️ Elemento <embed> não disponível ou sem src válido: {e_embed}", "INFO")

            return ""
        except Exception as e:
            log_message(f"⚠️ Não foi possível obter URL do PDF: {e}", "WARNING")
            return ""

    def prepare_window_open_capture(self, driver):
        """Sobrescreve window.open para capturar o link do relatório"""
        try:
            driver.execute_script("""
                try {
                    if (!window.__rpa_open_patched) {
                        window.__rpa_real_open = window.open;
                        window.open = function(url, name, specs) {
                            window.__rpa_last_open_url = url;
                            if (window.__rpa_real_open) {
                                return window.__rpa_real_open.call(this, url, name, specs);
                            }
                            return null;
                        };
                        window.__rpa_open_patched = true;
                    }
                    window.__rpa_last_open_url = null;
                } catch (e) {
                    window.__rpa_last_open_url = null;
                }
            """)
        except Exception as e:
            log_message(f"⚠️ Não foi possível preparar captura do window.open: {e}", "WARNING")

    def get_captured_window_open_url(self, driver, base_url: str) -> str:
        """Retorna o último link registrado pelo window.open sobrescrito"""
        try:
            captured_url = driver.execute_script("return window.__rpa_last_open_url || null;")
            if captured_url:
                absolute_url = urljoin(base_url, captured_url)
                log_message(f"✅ URL obtida via captura do window.open: {absolute_url}", "SUCCESS")
                return absolute_url
        except Exception as e:
            log_message(f"⚠️ Não foi possível recuperar URL capturada do window.open: {e}", "WARNING")
        return ""

    def download_pdf(self, driver, pdf_url: str, base_url: str, download_dir: str) -> str:
        """Realiza o download do PDF utilizando os cookies da sessão atual"""
        try:
            if not pdf_url:
                log_message("⚠️ URL do PDF vazia - download não realizado.", "WARNING")
                return ""

            parsed_url = urlparse(pdf_url)
            if not parsed_url.scheme:
                pdf_url = urljoin(base_url, pdf_url)
                parsed_url = urlparse(pdf_url)

            session = requests.Session()
            add_dict_to_cookiejar(session.cookies, {cookie['name']: cookie['value'] for cookie in driver.get_cookies()})

            log_message(f"📥 Baixando PDF diretamente da URL: {pdf_url}", "INFO")
            response = session.get(pdf_url, stream=True, timeout=60)
            response.raise_for_status()

            query_params = parse_qs(parsed_url.query)
            file_name = "relatorio.pdf"

            if "path" in query_params and query_params["path"]:
                raw_path = query_params["path"][0]
                decoded_path = unquote(raw_path)
                candidate_name = os.path.basename(decoded_path)
                if candidate_name:
                    file_name = candidate_name

            base_name, ext = os.path.splitext(file_name)
            if ext.lower() != ".pdf":
                file_name = f"{base_name}.pdf"

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = os.path.join(download_dir, file_name)
            if os.path.exists(file_path):
                file_path = os.path.join(download_dir, f"{base_name}_{timestamp}.pdf")

            with open(file_path, "wb") as pdf_file:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        pdf_file.write(chunk)

            log_message(f"✅ PDF baixado com sucesso em: {file_path}", "SUCCESS")
            return file_path
        except Exception as e:
            log_message(f"⚠️ Erro ao baixar PDF via HTTP: {e}", "WARNING")
            return ""

    def run(self, params: dict):
        username = params.get("username")
        password = params.get("password")
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode")
        excel_file = params.get("excel_file")
        data_tipo = params.get("data_tipo", "recepcao")
        cobrar_de = params.get("cobrar_de", "C")

        # Criar pasta de downloads específica para fatura mensal
        from pathlib import Path
        download_dir = os.path.join(os.getcwd(), "downloads", "fatura_mensal")
        Path(download_dir).mkdir(parents=True, exist_ok=True)
        log_message(f"📁 Pasta de downloads: {download_dir}", "INFO")
        
        url = os.getenv("SYSTEM_URL", "https://pathoweb.com.br/login/auth")
        parsed_url = urlparse(url)
        base_url = f"{parsed_url.scheme}://{parsed_url.netloc}" if parsed_url.scheme and parsed_url.netloc else "https://pathoweb.com.br"
        driver = BrowserFactory.create_chrome(download_dir=download_dir, headless=headless_mode)
        wait = WebDriverWait(driver, 15)

        try:
            log_message("Iniciando automação de Fatura Mensal...", "INFO")
            log_message(f"Tipo de data selecionado: {data_tipo}", "INFO")
            log_message(f"Tipo de cobrança selecionado: {cobrar_de}", "INFO")

            # Carregar dados do Excel
            if not excel_file or not os.path.exists(excel_file):
                messagebox.showerror("Erro", "Arquivo Excel não informado ou não encontrado.")
                return
            
            try:
                dados_excel = self.get_excel_data(excel_file)
            except Exception as e:
                messagebox.showerror("Erro", str(e))
                return
            
            if not dados_excel:
                messagebox.showerror("Erro", "Nenhum dado encontrado no arquivo Excel.")
                return
            
            log_message(f"✅ Carregados {len(dados_excel)} registros do Excel", "SUCCESS")

            # Login
            driver.get(url)
            wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
            driver.find_element(By.ID, "password").send_keys(password)
            driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

            log_message("Verificando se precisa navegar para módulo de faturamento...", "INFO")
            current_url = driver.current_url

            if current_url == "https://pathoweb.com.br/" or "trocarModulo" in current_url:
                log_message("Detectada tela de seleção de módulos - navegando para módulo de faturamento...", "INFO")
                try:
                    modulo_link = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=2']")))
                    modulo_link.click()
                    time.sleep(2)
                    log_message("✅ Navegação para módulo de faturamento realizada", "SUCCESS")
                except Exception as e:
                    log_message(f"⚠️ Erro ao navegar para módulo: {e}", "WARNING")
                    driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                    time.sleep(2)
                    log_message("🔄 Navegação direta para módulo realizada", "INFO")

            elif "moduloFaturamento" in current_url:
                log_message("✅ Já está no módulo de faturamento - pulando navegação", "SUCCESS")
            else:
                log_message(f"⚠️ URL inesperada detectada: {current_url}", "WARNING")
                driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                time.sleep(2)
                log_message("🔄 Navegação direta para módulo realizada (fallback)", "INFO")

            # Fechar modal se existir
            try:
                modal_close_button = driver.find_element(By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button")
                if modal_close_button.is_displayed():
                    modal_close_button.click()
                    time.sleep(1)
                    log_message("✅ Modal fechado", "SUCCESS")
            except Exception:
                log_message("ℹ️ Modal não encontrado ou já fechado", "INFO")

            # Acessar explicitamente a página do módulo de faturamento
            log_message("Acessando módulo de faturamento via URL...", "INFO")
            driver.get("https://pathoweb.com.br/moduloFaturamento/index")

            # Clicar no botão "Preparar exames para fatura"
            log_message("Clicando em 'Preparar exames para fatura'...", "INFO")
            try:
                preparar_btn = wait.until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "a.btn.btn-danger.chamadaAjax.setupAjax[data-url='/moduloFaturamento/preFaturamento']"
                )))
                preparar_btn.click()
                log_message("✅ Botão 'Preparar exames para fatura' clicado com sucesso", "SUCCESS")
            except Exception:
                try:
                    preparar_btn = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]"
                    )))
                    preparar_btn.click()
                    log_message("✅ Botão 'Preparar exames para fatura' clicado com sucesso (método alternativo)", "SUCCESS")
                except Exception as e:
                    log_message(f"❌ Erro ao clicar no botão 'Preparar exames para fatura': {e}", "ERROR")
                    raise Exception(f"Não foi possível clicar no botão: {e}")

            # Aguardar possível spinner/modal carregar
            try:
                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                log_message("✅ Modal de carregamento fechado", "INFO")
            except Exception:
                time.sleep(1)

            log_message("Tela de Pré Faturamento aberta.", "SUCCESS")
            
            # Aguardar a tela carregar completamente antes de processar
            log_message("Aguardando tela carregar completamente...", "INFO")
            time.sleep(5)  # Aguardar mais tempo para garantir que a tela carregou
            
            # Verificar se os campos principais estão presentes
            try:
                wait.until(EC.presence_of_element_located((By.ID, "cobrarDe")))
                wait.until(EC.presence_of_element_located((By.ID, "situacaoFaturamento")))
                wait.until(EC.presence_of_element_located((By.ID, "pessoaFaturamentoId")))
                wait.until(EC.presence_of_element_located((By.ID, "etapa")))
                log_message("✅ Campos principais detectados - tela pronta", "SUCCESS")
            except Exception as e:
                log_message(f"⚠️ Campos principais não encontrados: {e}", "WARNING")
                time.sleep(3)  # Aguardar mais um pouco

            # Processar cada linha do Excel
            resultados = []
            for i, dados in enumerate(dados_excel):
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break
                
                try:
                    log_message(f"➡️ Processando linha {i+1}: {dados['cliente']}", "INFO")
                    
                    # Configurar campo "Cobrar de"
                    try:
                        cobrar_de_select = wait.until(EC.presence_of_element_located((By.ID, "cobrarDe")))
                        select_cobrar = Select(cobrar_de_select)
                        select_cobrar.select_by_value(cobrar_de)
                        log_message(f"✅ Campo 'Cobrar de' configurado para: {cobrar_de}", "SUCCESS")
                        time.sleep(1)  # Aguardar processamento
                    except Exception as e:
                        log_message(f"⚠️ Erro ao configurar 'Cobrar de': {e}", "WARNING")
                    
                    # Configurar "Situação de faturamento" como "Não enviado"
                    try:
                        situacao_select = wait.until(EC.presence_of_element_located((By.ID, "situacaoFaturamento")))
                        select_situacao = Select(situacao_select)
                        select_situacao.select_by_value("A")  # "Não enviado"
                        log_message("✅ Situação de faturamento configurada como 'Não enviado'", "SUCCESS")
                        time.sleep(1)  # Aguardar processamento
                    except Exception as e:
                        log_message(f"⚠️ Erro ao configurar situação de faturamento: {e}", "WARNING")
                    
                    # Configurar "Empresa do faturamento" como "DAP"
                    try:
                        empresa_select = wait.until(EC.presence_of_element_located((By.ID, "pessoaFaturamentoId")))
                        select_empresa = Select(empresa_select)
                        select_empresa.select_by_value("43")  # "DAP - DIAGNOSTICO EM ANATOMIA PATOLOGICA"
                        log_message("✅ Empresa do faturamento configurada como 'DAP'", "SUCCESS")
                        time.sleep(1)  # Aguardar processamento
                    except Exception as e:
                        log_message(f"⚠️ Erro ao configurar empresa do faturamento: {e}", "WARNING")
                    
                    # Configurar "Etapa do exame" como vazio
                    try:
                        etapa_select = wait.until(EC.presence_of_element_located((By.ID, "etapa")))
                        select_etapa = Select(etapa_select)
                        select_etapa.select_by_value("")  # Valor vazio
                        log_message("✅ Etapa do exame configurada como vazio", "SUCCESS")
                        time.sleep(1)  # Aguardar processamento
                    except Exception as e:
                        log_message(f"⚠️ Erro ao configurar etapa do exame: {e}", "WARNING")
                    
                    # Configurar convênio ou procedência baseado na coluna B
                    if dados['tipo'].upper() == "CONVENIO":
                        try:
                            # Primeiro tentar com Select2 aberto (método mais direto)
                            if self.select_select2_option(driver, "convenioId", dados['cliente']):
                                log_message(f"✅ Convênio selecionado via Select2: {dados['cliente']}", "SUCCESS")
                            else:
                                # Se Select2 não funcionar, tentar métodos alternativos
                                convenio_select = wait.until(EC.presence_of_element_located((By.ID, "convenioId")))
                                option_value = self.find_option_by_text(driver, convenio_select, dados['cliente'])
                                if option_value:
                                    # Tentar seleção direta primeiro
                                    try:
                                        select_convenio = Select(convenio_select)
                                        select_convenio.select_by_value(option_value)
                                        log_message(f"✅ Convênio selecionado via Select: {dados['cliente']}", "SUCCESS")
                                        time.sleep(1)
                                    except Exception:
                                        # Se falhar, tentar com JavaScript para Select2
                                        try:
                                            driver.execute_script(f"$('#convenioId').val('{option_value}').trigger('change');")
                                            log_message(f"✅ Convênio selecionado via JavaScript: {dados['cliente']}", "SUCCESS")
                                            time.sleep(1)
                                        except Exception as e2:
                                            log_message(f"⚠️ Erro ao selecionar convênio via JavaScript: {e2}", "WARNING")
                                else:
                                    log_message(f"⚠️ Convênio não encontrado: {dados['cliente']}", "WARNING")
                        except Exception as e:
                            log_message(f"⚠️ Erro ao selecionar convênio: {e}", "WARNING")
                    
                    elif dados['tipo'].upper() == "PROCEDENCIA":
                        try:
                            # Primeiro tentar com Select2 aberto (método mais direto)
                            if self.select_select2_option(driver, "procedenciaId", dados['cliente']):
                                log_message(f"✅ Procedência selecionada via Select2: {dados['cliente']}", "SUCCESS")
                            else:
                                # Se Select2 não funcionar, tentar métodos alternativos
                                procedencia_select = wait.until(EC.presence_of_element_located((By.ID, "procedenciaId")))
                                option_value = self.find_option_by_text(driver, procedencia_select, dados['cliente'])
                                
                                if option_value:
                                    # Tentar seleção direta primeiro
                                    try:
                                        select_procedencia = Select(procedencia_select)
                                        select_procedencia.select_by_value(option_value)
                                        log_message(f"✅ Procedência selecionada via Select: {dados['cliente']}", "SUCCESS")
                                        time.sleep(1)
                                    except Exception:
                                        # Se falhar, tentar com JavaScript para Select2
                                        try:
                                            driver.execute_script(f"$('#procedenciaId').val('{option_value}').trigger('change');")
                                            log_message(f"✅ Procedência selecionada via JavaScript: {dados['cliente']}", "SUCCESS")
                                            time.sleep(1)
                                        except Exception as e2:
                                            log_message(f"⚠️ Erro ao selecionar procedência via JavaScript: {e2}", "WARNING")
                                else:
                                    log_message(f"⚠️ Procedência não encontrada: {dados['cliente']}", "WARNING")
                        except Exception as e:
                            log_message(f"⚠️ Erro ao selecionar procedência: {e}", "WARNING")
                    
                    # Configurar datas baseado na coluna C
                    if dados['data']:
                        try:
                            data_inicio, data_fim = self.parse_date_range(dados['data'])
                            
                            if data_tipo == "recepcao":
                                campo_data = wait.until(EC.presence_of_element_located((By.ID, "dataRecepcao")))
                                campo_data.clear()
                                campo_data.send_keys(f"{data_inicio} - {data_fim}")
                                log_message(f"✅ Data de recepção configurada: {data_inicio} - {data_fim}", "SUCCESS")
                                time.sleep(1)  # Aguardar processamento
                            else:  # liberacao
                                campo_data = wait.until(EC.presence_of_element_located((By.ID, "dataLiberacao")))
                                campo_data.clear()
                                campo_data.send_keys(f"{data_inicio} - {data_fim}")
                                log_message(f"✅ Data de liberação configurada: {data_inicio} - {data_fim}", "SUCCESS")
                                time.sleep(1)  # Aguardar processamento
                        except Exception as e:
                            log_message(f"⚠️ Erro ao configurar data: {e}", "WARNING")
                    
                    # Clicar no botão "Pesquisar" após configurar todos os campos
                    try:
                        pesquisar_btn = wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento")))
                        pesquisar_btn.click()
                        log_message("✅ Botão 'Pesquisar' clicado", "SUCCESS")
                        
                        # Aguardar carregamento dos resultados
                        try:
                            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                            log_message("🔄 Aguardando carregamento dos resultados...", "INFO")
                            WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                            log_message("✅ Resultados carregados", "SUCCESS")
                        except Exception:
                            time.sleep(2)  # Aguardar um tempo fixo se não encontrar spinner
                        
                        # Aguardar mais tempo após o modal fechar para garantir que a página processou
                        time.sleep(5)
                        
                        # Verificar se o modal realmente fechou antes de continuar
                        try:
                            spinner = driver.find_element(By.ID, "spinner")
                            if spinner.is_displayed():
                                log_message("⚠️ Modal ainda está visível, aguardando mais...", "WARNING")
                                WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                                time.sleep(3)
                        except:
                            log_message("✅ Modal confirmado como fechado", "SUCCESS")
                        
                        # PRIMEIRO: Desmarcar o checkbox "gerarArquivoTiss" ANTES de clicar no relatório
                        try:
                            checkbox = driver.find_element(By.ID, "gerarArquivoTiss")
                            if checkbox.is_selected():
                                # Tentar clicar normalmente primeiro
                                try:
                                    checkbox.click()
                                    log_message("✅ Checkbox 'gerarArquivoTiss' desmarcado ANTES do relatório", "SUCCESS")
                                except Exception:
                                    # Se falhar, usar JavaScript
                                    driver.execute_script("arguments[0].click();", checkbox)
                                    log_message("✅ Checkbox 'gerarArquivoTiss' desmarcado via JavaScript ANTES do relatório", "SUCCESS")
                                time.sleep(1)
                            else:
                                log_message("ℹ️ Checkbox 'gerarArquivoTiss' já estava desmarcado", "INFO")
                        except Exception as e_checkbox:
                            log_message(f"⚠️ Erro ao desmarcar checkbox antes do relatório: {e_checkbox}", "WARNING")
                        
                        # SEGUNDO: Clicar no botão "Relatório"
                        try:
                            self.prepare_window_open_capture(driver)
                            existing_downloads = set(os.listdir(download_dir)) if os.path.isdir(download_dir) else set()

                            # Tentar clicar com JavaScript se o clique normal falhar
                            try:
                                relatorio_btn = wait.until(EC.element_to_be_clickable((By.ID, "relatorioFaturamento")))
                                relatorio_btn.click()
                                log_message("✅ Botão 'Relatório' clicado", "SUCCESS")
                            except Exception as click_error:
                                log_message(f"⚠️ Clique normal falhou, tentando JavaScript: {click_error}", "WARNING")
                                driver.execute_script("document.getElementById('relatorioFaturamento').click();")
                                log_message("✅ Botão 'Relatório' clicado via JavaScript", "SUCCESS")
                            
                            # Aguardar modal do relatório aparecer e fechar
                            try:
                                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                                log_message("🔄 Aguardando geração do relatório...", "INFO")
                                WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                                log_message("✅ Relatório gerado", "SUCCESS")
                            except Exception:
                                time.sleep(3)  # Aguardar um tempo fixo se não encontrar spinner
                            
                            # Aguardar nova aba abrir com o PDF
                            log_message("🔄 Aguardando nova aba com PDF abrir...", "INFO")
                            time.sleep(3)  # Aguardar aba abrir
                            
                            # Salvar referência da aba original
                            original_window = driver.current_window_handle
                            log_message(f"📝 Aba original: {original_window}", "INFO")
                            
                            # Verificar se uma nova aba foi aberta
                            all_windows = driver.window_handles
                            log_message(f"🔍 Total de abas abertas: {len(all_windows)}", "INFO")
                            
                            pdf_current_url = ""
                            same_window_navigation = False
                            downloaded_file = ""

                            if len(all_windows) > 1:
                                # Trocar para a nova aba (PDF)
                                for window in all_windows:
                                    if window != original_window:
                                        driver.switch_to.window(window)
                                        log_message(f"✅ Trocado para nova aba com PDF", "SUCCESS")
                                        break
                                
                                # Aguardar página carregar completamente
                                time.sleep(3)
                                
                                # Tentar obter URL do PDF do DOM
                                pdf_current_url = self.get_pdf_url(driver, base_url)
                                
                                if pdf_current_url:
                                    log_message(f"📄 URL do PDF identificada: {pdf_current_url}", "SUCCESS")
                                    # Baixar PDF diretamente usando a URL
                                    downloaded_file = self.download_pdf(driver, pdf_current_url, base_url, download_dir)
                                    if downloaded_file:
                                        log_message(f"✅ PDF baixado diretamente: {downloaded_file}", "SUCCESS")
                                else:
                                    log_message("⚠️ Não foi possível obter URL do PDF. Tentando clicar no botão 'Abrir'...", "WARNING")
                                    # Fallback: tentar clicar no botão "Abrir"
                                    try:
                                        log_message("🔍 Procurando botão 'Abrir' com ID 'open-button'...", "INFO")
                                        
                                        # Tentar encontrar o botão por ID primeiro
                                        try:
                                            abrir_btn = WebDriverWait(driver, 10).until(
                                                EC.element_to_be_clickable((By.ID, "open-button"))
                                            )
                                            abrir_btn.click()
                                            log_message("✅ Botão 'Abrir' (ID: open-button) clicado", "SUCCESS")
                                            time.sleep(3)  # Aguardar download iniciar
                                        except Exception as e1:
                                            log_message(f"⚠️ Erro ao clicar por ID: {e1}", "WARNING")
                                            # Tentar por XPath com texto
                                            try:
                                                abrir_btn = WebDriverWait(driver, 5).until(
                                                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Abrir')]"))
                                                )
                                                abrir_btn.click()
                                                log_message("✅ Botão 'Abrir' clicado via XPath", "SUCCESS")
                                                time.sleep(3)
                                            except Exception as e2:
                                                log_message(f"⚠️ Erro ao clicar via XPath: {e2}", "WARNING")
                                                # Tentar JavaScript como último recurso
                                                try:
                                                    driver.execute_script("document.getElementById('open-button').click();")
                                                    log_message("✅ Botão 'Abrir' clicado via JavaScript", "SUCCESS")
                                                    time.sleep(3)
                                                except Exception as e3:
                                                    log_message(f"⚠️ Erro ao clicar via JavaScript: {e3}", "WARNING")
                                    except Exception as e_btn:
                                        log_message(f"⚠️ Erro geral ao processar botão 'Abrir': {e_btn}", "WARNING")
                                
                                # Aguardar download completar
                                log_message("📥 Aguardando download completar...", "INFO")
                                time.sleep(5)
                                
                                # Se ainda não baixou, tentar verificar se foi baixado pelo clique no botão
                                if not downloaded_file:
                                    downloaded_file = self.wait_for_download(download_dir, existing_downloads)
                                
                                # Fechar a aba do PDF
                                try:
                                    driver.close()
                                    log_message("✅ Aba do PDF fechada", "SUCCESS")
                                except Exception as e_close:
                                    log_message(f"⚠️ Erro ao fechar aba PDF: {e_close}", "WARNING")
                                
                                # Voltar para a aba original
                                driver.switch_to.window(original_window)
                                log_message("✅ Voltado para aba original", "SUCCESS")
                                time.sleep(1)
                            else:
                                log_message("⚠️ Nova aba não detectada, continuando...", "WARNING")
                                pdf_current_url = self.get_pdf_url(driver, base_url)
                                if pdf_current_url:
                                    same_window_navigation = True
                                    downloaded_file = self.download_pdf(driver, pdf_current_url, base_url, download_dir)
                                else:
                                    pdf_current_url = self.get_captured_window_open_url(driver, base_url)
                                    if pdf_current_url:
                                        downloaded_file = self.download_pdf(driver, pdf_current_url, base_url, download_dir)

                            # Se ainda não baixou, tentar mais uma vez com URL capturada
                            if not downloaded_file and not pdf_current_url:
                                pdf_current_url = self.get_captured_window_open_url(driver, base_url)
                                if pdf_current_url:
                                    downloaded_file = self.download_pdf(driver, pdf_current_url, base_url, download_dir)

                            # Último fallback: verificar se arquivo foi baixado
                            if not downloaded_file:
                                downloaded_file = self.wait_for_download(download_dir, existing_downloads)

                            if downloaded_file:
                                log_message(f"✅ PDF salvo em: {downloaded_file}", "SUCCESS")
                            else:
                                log_message("⚠️ Nenhum novo PDF detectado após a tentativa de download.", "WARNING")

                            if same_window_navigation:
                                try:
                                    log_message("🔄 Retornando para tela anterior após navegação na mesma aba...", "INFO")
                                    driver.back()
                                    wait.until(EC.presence_of_element_located((By.ID, "pesquisaFaturamento")))
                                    log_message("✅ Retorno para tela de Pré Faturamento concluído", "SUCCESS")
                                    time.sleep(2)
                                except Exception as back_error:
                                    log_message(f"⚠️ Erro ao retornar para tela anterior: {back_error}", "WARNING")
                            
                            # TERCEIRO: Clicar no botão "Situação faturamento para"
                            try:
                                log_message("🔄 Clicando em 'Situação faturamento para'...", "INFO")
                                
                                # Aguardar qualquer modal fechar
                                try:
                                    WebDriverWait(driver, 5).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                                except:
                                    pass
                                
                                time.sleep(2)
                                
                                # Clicar no botão
                                try:
                                    situacao_btn = wait.until(EC.element_to_be_clickable((By.ID, "executarMudancaSitFaturamento")))
                                    situacao_btn.click()
                                    log_message("✅ Botão 'Situação faturamento para' clicado", "SUCCESS")
                                except Exception:
                                    # Tentar com JavaScript
                                    driver.execute_script("document.getElementById('executarMudancaSitFaturamento').click();")
                                    log_message("✅ Botão 'Situação faturamento para' clicado via JavaScript", "SUCCESS")
                                
                                # Aguardar processamento
                                time.sleep(3)
                                
                                # Aguardar modal aparecer e fechar
                                try:
                                    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                                    log_message("🔄 Aguardando processamento de situação...", "INFO")
                                    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                                    log_message("✅ Situação de faturamento atualizada", "SUCCESS")
                                except Exception:
                                    log_message("ℹ️ Processamento de situação concluído", "INFO")
                                
                                time.sleep(2)
                                
                            except Exception as e_situacao:
                                log_message(f"⚠️ Erro ao clicar em 'Situação faturamento para': {e_situacao}", "WARNING")
                                
                        except Exception as e:
                            log_message(f"⚠️ Erro ao clicar no botão 'Relatório': {e}", "WARNING")
                            
                    except Exception as e:
                        log_message(f"⚠️ Erro ao clicar no botão 'Pesquisar': {e}", "WARNING")
                    
                    resultados.append({"linha": i+1, "cliente": dados['cliente'], "status": "sucesso"})
                    log_message(f"✅ Linha {i+1} processada com sucesso", "SUCCESS")
                    
                    # Aguardar mais tempo entre processamentos para garantir que tudo foi processado
                    log_message("Aguardando antes do próximo processamento...", "INFO")
                    time.sleep(5)
                    
                except Exception as e:
                    log_message(f"❌ Erro ao processar linha {i+1}: {e}", "ERROR")
                    resultados.append({"linha": i+1, "cliente": dados['cliente'], "status": "erro", "erro": str(e)})

            # Resumo final
            total = len(resultados)
            sucesso = [r for r in resultados if r["status"] == "sucesso"]
            erro = [r for r in resultados if r["status"] == "erro"]
            
            log_message("\nResumo do processamento:", "INFO")
            log_message(f"Total de linhas: {total}", "INFO")
            log_message(f"Processadas com sucesso: {len(sucesso)}", "SUCCESS")
            log_message(f"Erros: {len(erro)}", "ERROR")
            
            messagebox.showinfo("Sucesso",
                f"✅ Fatura Mensal processada com sucesso!\n\n"
                f"Total de linhas: {total}\n"
                f"Sucesso: {len(sucesso)}\n"
                f"Erros: {len(erro)}\n\n"
                f"Configurações aplicadas:\n"
                f"- Tipo de data: {data_tipo}\n"
                f"- Tipo de cobrança: {'Convênio' if cobrar_de == 'C' else 'Procedência'}"
            )

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            driver.quit()


def run(params: dict):
    module = FaturaMensalModule()
    module.run(params)
