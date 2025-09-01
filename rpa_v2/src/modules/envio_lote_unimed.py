import os
import time
import zipfile
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

class UnimedUploader(BaseModule):
    def __init__(self, username, password, driver=None, timeout=15, headless=False):
        super().__init__(nome="Envio Lote Unimed")
        self.username = username
        self.password = password
        self.timeout = timeout
        self.driver = driver
        self.wait = None
        self.headless = headless

    def inicializar_driver(self):
        if self.driver is None:
            log_message("Inicializando driver do Chrome para upload Unimed...", "INFO")
            self.driver = BrowserFactory.create_chrome(headless=self.headless)
        self.wait = WebDriverWait(self.driver, self.timeout)

    def fazer_login(self):
        log_message("Fazendo login no portal Unimed...", "INFO")
        self.driver.get("https://webmed.unimedlondrina.com.br/prestador/")
        campo_usuario = self.wait.until(EC.presence_of_element_located((By.ID, "operador")))
        campo_usuario.clear()
        campo_usuario.send_keys(self.username)
        campo_senha = self.driver.find_element(By.ID, "senha")
        campo_senha.clear()
        campo_senha.send_keys(self.password)
        botao_entrar = self.driver.find_element(By.ID, "entrar")
        botao_entrar.click()
        time.sleep(2.5)

    def acessar_url_pos_login(self, url_pos_login):
        log_message("Acessando página de upload TISS...", "INFO")
        self.driver.get(url_pos_login)

    def selecionar_versao_upload(self):
        log_message("Selecionando versão 4.01.00 para upload...", "INFO")
        select_element = self.wait.until(EC.presence_of_element_located((By.ID, "versao")))
        select_obj = Select(select_element)
        select_obj.select_by_value("4.01.00")

    def selecionar_arquivo_upload(self, caminho_arquivo):
        log_message(f"Selecionando arquivo para upload: {caminho_arquivo}", "INFO")
        input_arquivo = self.wait.until(EC.presence_of_element_located((By.ID, "arquivo")))
        input_arquivo.send_keys(caminho_arquivo)

    def clicar_enviar_upload(self):
        log_message("Enviando arquivo para Unimed...", "INFO")
        botao_enviar = self.wait.until(EC.element_to_be_clickable((By.ID, "enviar2")))
        botao_enviar.click()
        time.sleep(2)
        try:
            form_erro = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//form[contains(@action, 'relatorioErroXml.php')]"))
            )
            visualizar_link = form_erro.find_element(By.XPATH, ".//a[contains(text(), 'Visualizar erros')]")
            visualizar_link.click()
            log_message("Arquivo processado com erros. Baixando relatório de erros (PDF)...", "ERROR")
            time.sleep(1)
        except Exception:
            log_message("Arquivo enviado e processado sem erros.", "SUCCESS")

    def fechar(self):
        if self.driver:
            log_message("Fechando navegador do upload Unimed.", "INFO")
            self.driver.quit()

class XMLGeneratorAutomation(BaseModule):
    def __init__(self, username, password, timeout=15, pasta_download=None, fechar_em_erro=False, headless=False):
        super().__init__(nome="Geração e Envio XML Unimed")
        self.username = username
        self.password = password
        self.timeout = timeout
        self.driver = None
        self.wait = None
        self.arquivos_extraidos = []
        self.fechar_em_erro = fechar_em_erro
        if pasta_download is None:
            self.pasta_download = os.path.join(os.getcwd(), "downloads")
        else:
            self.pasta_download = pasta_download
        Path(self.pasta_download).mkdir(parents=True, exist_ok=True)

        self.headless_mode = headless

    def inicializar_driver(self):
        log_message("Inicializando driver do Chrome para Pathoweb...", "INFO")
        self.driver = BrowserFactory.create_chrome(download_dir=self.pasta_download, headless=self.headless_mode)
        self.wait = WebDriverWait(self.driver, self.timeout)

    def fazer_login(self):
        log_message("Fazendo login no Pathoweb...", "INFO")
        self.driver.get("https://pathoweb.com.br/login/auth")
        campo_usuario = self.wait.until(EC.presence_of_element_located((By.ID, "username")))
        campo_usuario.send_keys(self.username)
        campo_senha = self.driver.find_element(By.ID, "password")
        campo_senha.send_keys(self.password)
        botao_login = self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        botao_login.click()
        self.wait.until(lambda driver: "login" not in driver.current_url.lower())

    def acessar_modulo_faturamento(self):
        log_message("Acessando módulo de faturamento...", "INFO")
        link_faturamento = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=2']"))
        )
        link_faturamento.click()

    def fechar_modal_se_necessario(self):
        time.sleep(4)
        try:
            modal_close_button = self.driver.find_element(
                By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button"
            )
            if modal_close_button.is_displayed():
                modal_close_button.click()
                time.sleep(1)
        except NoSuchElementException:
            pass

    def acessar_preparar_exames_para_fatura(self):
        log_message("Acessando tela 'Preparar exames para fatura'...", "INFO")
        link_preparar = self.wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]"))
        )
        link_preparar.click()
        time.sleep(2)

    def configurar_filtro_convenio_unimed(self):
        log_message("Selecionando convênio UNIMED (LONDRINA)...", "INFO")
        select2_container = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-selection[aria-labelledby*='convenioId']"))
        )
        select2_container.click()
        time.sleep(1)
        opcao_unimed = self.wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//li[contains(@class, 'select2-results__option') and text()='UNIMED (LONDRINA)']"))
        )
        opcao_unimed.click()
        time.sleep(1)

    def configurar_filtro_conferido_online(self):
        log_message("Selecionando filtro 'Conferido Online'...", "INFO")
        select_element = self.wait.until(EC.presence_of_element_located((By.ID, "conferido")))
        select_conferido = Select(select_element)
        select_conferido.select_by_value("O")
        time.sleep(1)

    def executar_pesquisa_faturamento(self):
        log_message("Executando pesquisa de faturamento...", "INFO")
        botao_pesquisar = self.wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento")))
        botao_pesquisar.click()
        time.sleep(3)

    def aguardar_finalizacao_pesquisa(self):
        log_message("Aguardando finalização da pesquisa...", "INFO")
        tempo_maximo = time.time() + 60
        while time.time() < tempo_maximo:
            try:
                modal_carregando = self.driver.find_element(By.XPATH,
                                                            "//div[contains(@class,'modal-body') and contains(., 'Carregando')]")
                if modal_carregando.is_displayed():
                    time.sleep(1)
                else:
                    return
            except Exception:
                return

    def clicar_botao_situacao_faturamento(self):
        log_message("Clicando para baixar o lote XML...", "INFO")
        botao = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.btn.btn-danger[onclick*='modalFaturamento']"))
        )
        botao.click()

    def aguardar_download_completar(self, timeout_download=150):
        log_message("Aguardando download do arquivo XML/ZIP...", "INFO")
        arquivos_antes = set(os.listdir(self.pasta_download))
        tempo_limite = time.time() + timeout_download
        while time.time() < tempo_limite:
            arquivos_agora = set(os.listdir(self.pasta_download))
            novos_arquivos = arquivos_agora - arquivos_antes
            for arquivo in novos_arquivos:
                if arquivo.endswith(('.zip', '.xml', '.ZIP', '.XML')) and not arquivo.endswith('.crdownload'):
                    log_message(f"Arquivo baixado: {arquivo}", "SUCCESS")
                    return os.path.join(self.pasta_download, arquivo)
            time.sleep(2)
        log_message("Timeout ao aguardar download do arquivo.", "ERROR")
        return None

    def extrair_arquivo_zip(self, caminho_zip):
        if not caminho_zip or not os.path.exists(caminho_zip):
            log_message("Arquivo ZIP não encontrado para extração.", "ERROR")
            return []
        if not caminho_zip.lower().endswith('.zip'):
            return [caminho_zip]
        pasta_extracao = os.path.dirname(caminho_zip)
        arquivos_extraidos = []
        try:
            with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
                for arquivo in zip_ref.namelist():
                    zip_ref.extract(arquivo, pasta_extracao)
                    arquivos_extraidos.append(os.path.join(pasta_extracao, arquivo))
            log_message(f"Arquivos extraídos: {arquivos_extraidos}", "SUCCESS")
            os.remove(caminho_zip)
        except Exception as e:
            log_message(f"Erro ao extrair ZIP: {e}", "ERROR")
        return arquivos_extraidos

    def processar_download_completo(self):
        self.clicar_botao_situacao_faturamento()
        arquivo_baixado = self.aguardar_download_completar()
        if not arquivo_baixado:
            log_message("Nenhum arquivo baixado para processar.", "ERROR")
            return []
        arquivos_finais = self.extrair_arquivo_zip(arquivo_baixado)
        return arquivos_finais

    def configurar_filtros_e_pesquisar(self):
        self.configurar_filtro_convenio_unimed()
        time.sleep(1)
        self.configurar_filtro_conferido_online()
        self.executar_pesquisa_faturamento()
        self.aguardar_finalizacao_pesquisa()
        time.sleep(1)

    def enviar_para_unimed(self, arquivos_extraidos, unimed_user, unimed_pass):
        uploader = UnimedUploader(unimed_user, unimed_pass, self.driver)
        uploader.inicializar_driver()
        uploader.fazer_login()
        uploader.acessar_url_pos_login("https://webmed.unimedlondrina.com.br/prestador/uploadTiss.php")
        uploader.selecionar_versao_upload()
        for arquivo in arquivos_extraidos:
            if arquivo.lower().endswith(".xml"):
                uploader.selecionar_arquivo_upload(arquivo)
                uploader.clicar_enviar_upload()
                log_message(f"Arquivo {arquivo} enviado para Unimed.", "SUCCESS")
        #uploader.fechar()

    def executar_processo_completo_login_navegacao(self, unimed_user, unimed_pass, cancel_flag=None):
        try:
            log_message("Iniciando automação de envio de lote Unimed...", "INFO")
            self.inicializar_driver()
            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada pelo usuário.", "WARNING")
                self.fechar_navegador()
                return False
            self.fazer_login()
            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada pelo usuário.", "WARNING")
                self.fechar_navegador()
                return False

            log_message("Verificando se precisa navegar para módulo de faturamento...", "INFO")
            current_url = self.driver.current_url

            if current_url == "https://pathoweb.com.br/" or "trocarModulo" in current_url:
                log_message("Detectada tela de seleção de módulos - navegando para módulo de faturamento...", "INFO")
                try:
                    modulo_link = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=2']")))
                    modulo_link.click()
                    time.sleep(2)
                    log_message("✅ Navegação para módulo de faturamento realizada", "SUCCESS")
                except Exception as e:
                    log_message(f"⚠️ Erro ao navegar para módulo: {e}", "WARNING")
                    # Tentar navegar diretamente pela URL como fallback
                    self.driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                    time.sleep(2)
                    log_message("🔄 Navegação direta para módulo realizada", "INFO")

            elif "moduloFaturamento" in current_url:
                log_message("✅ Já está no módulo de faturamento - pulando navegação", "SUCCESS")
            else:
                log_message(f"⚠️ URL inesperada detectada: {current_url}", "WARNING")
                # Tentar navegar diretamente como fallback
                self.driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                time.sleep(2)
                log_message("🔄 Navegação direta para módulo realizada (fallback)", "INFO")

            self.fechar_modal_se_necessario()
            self.acessar_preparar_exames_para_fatura()
            self.configurar_filtros_e_pesquisar()
            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada pelo usuário.", "WARNING")
                self.fechar_navegador()
                return False
            arquivos_extraidos = self.processar_download_completo()
            if not arquivos_extraidos:
                log_message("Nenhum arquivo extraído para envio.", "ERROR")
                self.fechar_navegador()
                return False
            self.arquivos_extraidos = arquivos_extraidos
            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada pelo usuário.", "WARNING")
                self.fechar_navegador()
                return False
            self.enviar_para_unimed(arquivos_extraidos, unimed_user, unimed_pass)
            log_message("Processo concluído com sucesso!", "SUCCESS")
            self.fechar_navegador()
            return True
        except Exception as e:
            log_message(f"Erro durante a automação: {e}", "ERROR")
            self.fechar_navegador()
            return False

    def fechar_navegador(self):
        if self.driver:
            log_message("Fechando navegador do Pathoweb.", "INFO")
            self.driver.quit()

def run(params):
    """
    Parâmetros esperados em params:
    - username: usuário do Pathoweb
    - password: senha do Pathoweb
    - unimed_user: usuário Unimed (opcional, pode fixar se preferir)
    - unimed_pass: senha Unimed (opcional, pode fixar se preferir)
    - pasta_download: pasta para salvar arquivos (opcional)
    """
    username = params.get("username")
    password = params.get("password")
    unimed_user = params.get("unimed_user")
    unimed_pass = params.get("unimed_pass")

    if not unimed_user or not unimed_pass:
        log_message("Credenciais da Unimed não fornecidas!", "ERROR")
        return False

    pasta_download = params.get("pasta_download", os.path.join(os.getcwd(), "xml"))
    cancel_flag = params.get("cancel_flag")
    headless_mode = params.get("headless_mode", False)
    module = XMLGeneratorAutomation(username, password, pasta_download=pasta_download, headless=headless_mode)
    sucesso = module.executar_processo_completo_login_navegacao(unimed_user, unimed_pass, cancel_flag=cancel_flag)
    if not sucesso:
        log_message("❌ Falha na automação de envio de lote Unimed.", "ERROR")