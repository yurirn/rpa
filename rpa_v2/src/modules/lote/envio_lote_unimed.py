import os
import time
import zipfile
import traceback
import re
from datetime import datetime
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
        # Timeout maior em headless para evitar crashes
        self.timeout = 60 if headless else timeout
        self.driver = driver
        self.wait = None
        self.headless = headless

    def inicializar_driver(self):
        if self.driver is None:
            log_message("Inicializando driver do Chrome para upload Unimed...", "INFO")
            self.driver = BrowserFactory.create_chrome(headless=self.headless)
            # Configurar timeouts do driver
            self.driver.set_page_load_timeout(120)
            self.driver.set_script_timeout(60)
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
        log_message(f"Selecionando versão 4.02.00 para upload...", "INFO")
        select_element = self.wait.until(EC.presence_of_element_located((By.ID, "versao")))
        select_obj = Select(select_element)
        select_obj.select_by_value("4.02.00")

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
        # Timeout maior em headless para evitar crashes
        self.timeout = 60 if headless else timeout
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
        # Configurar timeouts do driver para evitar GetHandleVerifier errors
        self.driver.set_page_load_timeout(120)
        self.driver.set_script_timeout(60)
        self.wait = WebDriverWait(self.driver, self.timeout)

    def fazer_login(self):
        log_message("Fazendo login no Pathoweb...", "INFO")
        self.driver.get("https://dap.pathoweb.com.br/login/auth")
        self.wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(self.username)
        self.driver.find_element(By.ID, "password").send_keys(self.password)
        self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
        time.sleep(0.5)
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
        
        # Tentar múltiplas vezes se necessário
        max_tentativas = 3
        for tentativa in range(1, max_tentativas + 1):
            try:
                # Verificar se o driver ainda está vivo
                try:
                    _ = self.driver.current_url
                except Exception as e:
                    raise Exception(f"Driver perdeu conexão: {e}")
                
                # Aguardar mais tempo em headless
                tempo_espera = 2 if self.headless_mode else 1
                time.sleep(tempo_espera)
                
                # Aguardar especificamente pelo select2 do convênio
                log_message(f"Aguardando elemento do convênio (tentativa {tentativa}/{max_tentativas})...", "INFO")
                timeout_wait = 30 if self.headless_mode else 20
                select2_container = WebDriverWait(self.driver, timeout_wait).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-selection[aria-labelledby*='convenioId']"))
                )
                
                # Scroll até o elemento para garantir que está visível
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select2_container)
                time.sleep(0.5)
                
                # Clicar no select2
                select2_container.click()
                time.sleep(1.5 if self.headless_mode else 1)
                
                # Aguardar e selecionar a opção UNIMED
                opcao_unimed = WebDriverWait(self.driver, 15).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//li[contains(@class, 'select2-results__option') and text()='UNIMED (LONDRINA)']"))
                )
                opcao_unimed.click()
                time.sleep(1.5 if self.headless_mode else 1)
                
                log_message("✅ Convênio UNIMED selecionado com sucesso!", "SUCCESS")
                return
                
            except Exception as e:
                log_message(f"⚠️ Tentativa {tentativa} falhou: {str(e)}", "WARNING")
                if tentativa < max_tentativas:
                    log_message("🔄 Tentando novamente...", "INFO")
                    time.sleep(3 if self.headless_mode else 2)
                else:
                    log_message("❌ Não foi possível selecionar o convênio após múltiplas tentativas", "ERROR")
                    raise

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
        # Timeout maior em headless
        timeout_pesquisa = 120 if self.headless_mode else 60
        tempo_maximo = time.time() + timeout_pesquisa
        
        while time.time() < tempo_maximo:
            try:
                # Verificar se o driver ainda está vivo
                _ = self.driver.current_url
                
                modal_carregando = self.driver.find_element(By.XPATH,
                                                            "//div[contains(@class,'modal-body') and contains(., 'Carregando')]")
                if modal_carregando.is_displayed():
                    time.sleep(2 if self.headless_mode else 1)
                else:
                    log_message("✅ Modal fechado, pesquisa finalizada", "SUCCESS")
                    return
            except Exception:
                # Modal não encontrado ou já fechou
                log_message("✅ Pesquisa finalizada (modal não encontrado)", "SUCCESS")
                return
        
        log_message("⚠️ Timeout ao aguardar finalização da pesquisa", "WARNING")

    def clicar_botao_situacao_faturamento(self):
        log_message("Clicando para baixar o lote XML...", "INFO")
        
        try:
            # Aguardar botão estar presente e visível
            botao = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a.btn.btn-danger[onclick*='modalFaturamento']"))
            )
            
            # Scroll até o botão para garantir que está visível
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)
            time.sleep(0.5)
            
            # Clicar no botão
            botao.click()
            log_message("✅ Botão de download clicado com sucesso", "SUCCESS")
            
            # Aguardar um momento para o download iniciar
            time.sleep(2)
            
        except Exception as e:
            log_message(f"❌ Erro ao clicar no botão de download: {e}", "ERROR")
            raise

    def aguardar_download_completar(self, timeout_download=None):
        # Timeout maior em headless
        if timeout_download is None:
            timeout_download = 180 if self.headless_mode else 60
            
        log_message(f"Aguardando download do arquivo XML/ZIP (timeout: {timeout_download}s)...", "INFO")
        log_message(f"📁 Pasta de download: {self.pasta_download}", "INFO")
        
        arquivos_antes = set(os.listdir(self.pasta_download))
        log_message(f"📋 Arquivos antes do download: {len(arquivos_antes)} arquivo(s)", "INFO")
        
        tempo_limite = time.time() + timeout_download
        tentativa = 0
        
        while time.time() < tempo_limite:
            try:
                # Verificar se o driver ainda está vivo
                _ = self.driver.current_url
            except Exception as e:
                log_message(f"⚠️ Driver pode ter crashado durante download: {e}", "WARNING")
            
            tentativa += 1
            arquivos_agora = set(os.listdir(self.pasta_download))
            novos_arquivos = arquivos_agora - arquivos_antes
            
            if tentativa % 5 == 0:  # Log a cada 10 segundos (5 tentativas * 2 segundos)
                log_message(f"⏳ Aguardando download... ({tentativa * 2}s / {timeout_download}s)", "INFO")
            
            # Verificar se há arquivos .crdownload (download em andamento)
            arquivos_em_download = [f for f in arquivos_agora if f.endswith('.crdownload')]
            if arquivos_em_download:
                log_message(f"📥 Download em andamento: {arquivos_em_download[0]}", "INFO")
            
            for arquivo in novos_arquivos:
                if arquivo.endswith(('.zip', '.xml', '.ZIP', '.XML')) and not arquivo.endswith('.crdownload'):
                    log_message(f"✅ Arquivo baixado: {arquivo}", "SUCCESS")
                    return os.path.join(self.pasta_download, arquivo)
            
            time.sleep(2)
        
        log_message(f"❌ Timeout ao aguardar download após {timeout_download}s", "ERROR")
        log_message(f"📋 Arquivos na pasta agora: {os.listdir(self.pasta_download)}", "ERROR")
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
        try:
            log_message("→ Clicando no botão de download do lote...", "INFO")
            self.clicar_botao_situacao_faturamento()
            
            log_message("→ Aguardando download completar...", "INFO")
            arquivo_baixado = self.aguardar_download_completar()
            
            if not arquivo_baixado:
                log_message("❌ Nenhum arquivo baixado para processar.", "ERROR")
                return []
            
            log_message(f"✅ Arquivo baixado: {os.path.basename(arquivo_baixado)}", "SUCCESS")
            
            log_message("→ Extraindo arquivo ZIP (se necessário)...", "INFO")
            arquivos_finais = self.extrair_arquivo_zip(arquivo_baixado)
            
            log_message(f"✅ {len(arquivos_finais)} arquivo(s) pronto(s) para envio", "SUCCESS")
            return arquivos_finais
        except Exception as e:
            erro_download = traceback.format_exc()
            log_message(f"❌ Erro ao processar download:", "ERROR")
            log_message(f"Detalhes: {str(e)}", "ERROR")
            log_message(f"Stack trace:\n{erro_download}", "ERROR")
            raise

    def configurar_filtros_e_pesquisar(self):
        try:
            # Limpar campos de pesquisa antes de configurar filtros
            log_message("🧹 Limpando campos de pesquisa...", "INFO")
            try:
                # Limpar campo de número de exame
                campo_exame = self.driver.find_element(By.ID, "numeroExame")
                campo_exame.clear()
                
                # Limpar campo de número de guia
                campo_guia = self.driver.find_element(By.ID, "numeroGuia")
                campo_guia.clear()
                
                log_message("✅ Campos limpos com sucesso", "SUCCESS")
            except Exception as e:
                log_message(f"⚠️ Aviso ao limpar campos: {e}", "WARNING")
            
            time.sleep(0.5)

            self.configurar_filtro_convenio_unimed()
            time.sleep(1)

            self.configurar_filtro_conferido_online()

            self.executar_pesquisa_faturamento()

            self.aguardar_finalizacao_pesquisa()
            time.sleep(1)
            
            log_message("✅ Filtros configurados e pesquisa concluída", "SUCCESS")
        except Exception as e:
            erro_filtros = traceback.format_exc()
            log_message(f"❌ Erro ao configurar filtros:", "ERROR")
            log_message(f"Detalhes: {str(e)}", "ERROR")
            log_message(f"Stack trace:\n{erro_filtros}", "ERROR")
            raise

    def enviar_para_unimed(self, arquivos_extraidos, unimed_user, unimed_pass):
        try:
            log_message("Iniciando processo de upload para Unimed...", "INFO")
            uploader = UnimedUploader(unimed_user, unimed_pass, self.driver)
            
            log_message("Inicializando driver para Unimed...", "INFO")
            uploader.inicializar_driver()
            
            log_message("Fazendo login na Unimed...", "INFO")
            uploader.fazer_login()
            
            log_message("Acessando página de upload TISS...", "INFO")
            uploader.acessar_url_pos_login("https://webmed.unimedlondrina.com.br/prestador/uploadTiss.php")
            
            log_message("Selecionando versão do upload...", "INFO")
            uploader.selecionar_versao_upload()
            
            for idx, arquivo in enumerate(arquivos_extraidos, 1):
                if arquivo.lower().endswith(".xml"):
                    log_message(f"Enviando arquivo {idx}/{len(arquivos_extraidos)}: {os.path.basename(arquivo)}", "INFO")
                    uploader.selecionar_arquivo_upload(arquivo)
                    uploader.clicar_enviar_upload()
                    log_message(f"Arquivo {arquivo} enviado para Unimed.", "SUCCESS")
            
            log_message("Todos os arquivos foram enviados com sucesso!", "SUCCESS")
            #uploader.fechar()
        except Exception as e:
            erro_upload = traceback.format_exc()
            log_message(f"❌ Erro durante upload para Unimed:", "ERROR")
            log_message(f"Tipo: {type(e).__name__}", "ERROR")
            log_message(f"Mensagem: {str(e)}", "ERROR")
            log_message(f"Stack trace:\n{erro_upload}", "ERROR")
            raise

    def verificar_carregamento_pagina(self, max_tentativas=3):
        """Verifica se a página de faturamento carregou corretamente."""
        for tentativa in range(1, max_tentativas + 1):
            try:
                log_message(f"🔍 Verificando carregamento da página (tentativa {tentativa}/{max_tentativas})...", "INFO")
                
                # Aguardar um pouco para a página começar a carregar
                time.sleep(2)
                
                # Verificar se a página não está em branco
                body_text = self.driver.execute_script("return document.body.innerText;")
                if not body_text or len(body_text.strip()) < 50:
                    log_message("⚠️ Página parece estar em branco, tentando recarregar...", "WARNING")
                    self.driver.refresh()
                    time.sleep(3)
                    continue
                
                # Tentar clicar na aba "Pré faturamento e faturar" se existir
                try:
                    log_message("🔍 Procurando pela aba 'Pré faturamento e faturar'...", "INFO")
                    aba_faturamento = self.driver.find_element(
                        By.XPATH, "//a[contains(text(), 'Pré faturamento e faturar') or contains(text(), 'faturamento')]"
                    )
                    if aba_faturamento.is_displayed():
                        log_message("✅ Aba encontrada, clicando...", "INFO")
                        aba_faturamento.click()
                        time.sleep(2)
                except Exception as e:
                    log_message(f"ℹ️ Aba não encontrada ou já está selecionada: {e}", "INFO")
                
                # Verificar se existe o formulário de pesquisa
                form_exists = self.driver.execute_script("""
                    return document.querySelector('#pesquisaFaturamento') !== null;
                """)
                
                if form_exists:
                    log_message("✅ Página de faturamento carregou corretamente!", "SUCCESS")
                    return True
                else:
                    log_message("⚠️ Formulário de pesquisa não encontrado, recarregando...", "WARNING")
                    self.driver.refresh()
                    time.sleep(3)
                    
            except Exception as e:
                log_message(f"⚠️ Erro ao verificar carregamento: {e}", "WARNING")
                if tentativa < max_tentativas:
                    log_message("🔄 Tentando recarregar a página...", "INFO")
                    self.driver.refresh()
                    time.sleep(3)
        
        log_message("❌ Página não carregou corretamente após múltiplas tentativas", "ERROR")
        return False

    def executar_processo_completo_sem_login(self, unimed_user, unimed_pass, cancel_flag=None, total_exames_lote=0):
        """
        Executa o processo completo de geração e envio do XML reutilizando o driver já aberto.
        Não faz login novamente, assume que o driver já está autenticado no Pathoweb.
        """
        try:
            log_message("🔄 Continuando no navegador já autenticado...", "INFO")
            
            # O driver e wait já foram configurados externamente
            if not self.driver or not self.wait:
                raise Exception("Driver ou Wait não foram configurados!")
            
            log_message("Verificando URL atual...", "INFO")
            current_url = self.driver.current_url
            log_message(f"URL atual: {current_url}", "INFO")
            
            # Verificar se está no módulo de faturamento
            if "moduloFaturamento" not in current_url:
                log_message("⚠️ Não está no módulo de faturamento. Navegando...", "WARNING")
                self.driver.get("https://dap.pathoweb.com.br/moduloFaturamento/index")
                time.sleep(3)
            else:
                log_message("✅ Já está no módulo de faturamento", "SUCCESS")
            
            # Fechar modal se necessário
            self.fechar_modal_se_necessario()
            
            # Verificar se já está na página de faturamento
            current_url = self.driver.current_url
            log_message(f"🔍 Verificando URL para navegação: {current_url}", "INFO")
            
            # Verificar se já tem o formulário de pesquisa carregado (sinal que já está na página)
            formulario_ja_presente = False
            try:
                formulario_ja_presente = self.driver.execute_script("""
                    return document.querySelector('#pesquisaFaturamento') !== null;
                """)
            except:
                pass
            
            if formulario_ja_presente and ("faturamento" in current_url.lower() or "preFaturamento" in current_url):
                log_message("✅ Já está na página de preparação de exames com formulário carregado, pulando navegação...", "SUCCESS")
            else:
                log_message("Navegando para 'Preparar exames para fatura'...", "INFO")
                try:
                    # Tentar clicar no link em vez de navegar diretamente (melhor para apps AJAX)
                    link_preparar = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]"))
                    )
                    link_preparar.click()
                    log_message("✅ Link clicado com sucesso", "SUCCESS")
                    time.sleep(3)
                except Exception as e:
                    log_message(f"⚠️ Não foi possível clicar no link: {e}", "WARNING")
                    log_message("Tentando navegação direta como fallback...", "INFO")
                    self.driver.get("https://dap.pathoweb.com.br/moduloFaturamento/faturamento")
                    time.sleep(2)
                
                # Verificar se a página carregou corretamente
                if not self.verificar_carregamento_pagina():
                    raise Exception("Página de faturamento não carregou corretamente após múltiplas tentativas")
            
            log_message("Configurando filtros e executando pesquisa...", "INFO")
            self.configurar_filtros_e_pesquisar()

            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada pelo usuário.", "WARNING")
                return False
            
            arquivos_extraidos = self.processar_download_completo()
            if not arquivos_extraidos:
                log_message("Nenhum arquivo extraído para envio.", "ERROR")
                return False
            
            self.arquivos_extraidos = arquivos_extraidos
            
            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada pelo usuário.", "WARNING")
                return False
            
            self.enviar_para_unimed(arquivos_extraidos, unimed_user, unimed_pass)
            log_message("Processo concluído com sucesso!", "SUCCESS")
            return True
            
        except Exception as e:
            erro_completo = traceback.format_exc()
            log_message(f"❌ Erro durante a automação de envio Unimed:", "ERROR")
            log_message(f"Tipo do erro: {type(e).__name__}", "ERROR")
            log_message(f"Mensagem: {str(e)}", "ERROR")
            log_message(f"Stack trace completo:\n{erro_completo}", "ERROR")
            return False

    def executar_processo_completo_login_navegacao(self, unimed_user, unimed_pass, cancel_flag=None, total_exames_lote=0):
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

            if current_url == "https://dap.pathoweb.com.br/" or "trocarModulo" in current_url:
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
                    self.driver.get("https://dap.pathoweb.com.br/moduloFaturamento/index")
                    time.sleep(2)
                    log_message("🔄 Navegação direta para módulo realizada", "INFO")

            elif "moduloFaturamento" in current_url:
                log_message("✅ Já está no módulo de faturamento - pulando navegação", "SUCCESS")
            else:
                log_message(f"⚠️ URL inesperada detectada: {current_url}", "WARNING")
                # Tentar navegar diretamente como fallback
                self.driver.get("https://dap.pathoweb.com.br/moduloFaturamento/index")
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
            erro_completo = traceback.format_exc()
            log_message(f"❌ Erro durante a automação de envio Unimed:", "ERROR")
            log_message(f"Tipo do erro: {type(e).__name__}", "ERROR")
            log_message(f"Mensagem: {str(e)}", "ERROR")
            log_message(f"Stack trace completo:\n{erro_completo}", "ERROR")
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