import os
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from dotenv import load_dotenv

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

load_dotenv()


class BaixaLoteModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Baixa de Lote")
        self.driver = None
        self.wait = None

    def inicializar_driver(self, headless_mode=False):
        """Inicializa o driver do navegador"""
        log_message("Inicializando navegador...", "INFO")
        self.driver = BrowserFactory.create_chrome(headless=headless_mode)
        self.wait = WebDriverWait(self.driver, 15)
        log_message("✅ Navegador inicializado", "SUCCESS")

    def fazer_login(self, username, password):
        """Realiza login no sistema"""
        try:
            url = os.getenv("SYSTEM_URL", "https://dap.pathoweb.com.br/login/auth")
            log_message(f"Acessando {url}...", "INFO")
            self.driver.get(url)

            log_message("Preenchendo credenciais...", "INFO")
            self.wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
            self.driver.find_element(By.ID, "password").send_keys(password)
            self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

            log_message("✅ Login realizado com sucesso", "SUCCESS")
            time.sleep(2)
        except Exception as e:
            log_message(f"❌ Erro ao fazer login: {e}", "ERROR")
            raise

    def navegar_para_modulo_financeiro(self):
        """Navega para o módulo financeiro"""
        try:
            log_message("Navegando para módulo financeiro...", "INFO")
            current_url = self.driver.current_url

            if current_url == "https://dap.pathoweb.com.br/" or "trocarModulo" in current_url:
                log_message("Detectada tela de seleção de módulos...", "INFO")
                try:
                    modulo_link = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=4']"))
                    )
                    modulo_link.click()
                    time.sleep(2)
                    log_message("✅ Navegação para módulo financeiro realizada", "SUCCESS")
                except Exception as e:
                    log_message(f"⚠️ Erro ao navegar para módulo: {e}", "WARNING")
                    self.driver.get("https://dap.pathoweb.com.br/moduloFinanceiro/index")
                    time.sleep(2)

            elif "moduloFinanceiro" in current_url:
                log_message("✅ Já está no módulo financeiro", "SUCCESS")
            else:
                log_message(f"⚠️ URL inesperada: {current_url}", "WARNING")
                self.driver.get("https://dap.pathoweb.com.br/moduloFinanceiro/index")
                time.sleep(2)

        except Exception as e:
            log_message(f"❌ Erro ao navegar para módulo financeiro: {e}", "ERROR")
            raise

    def fechar_modal_se_necessario(self):
        """Fecha modal se estiver aberto"""
        try:
            modal_close_button = self.driver.find_element(
                By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button"
            )
            if modal_close_button.is_displayed():
                modal_close_button.click()
                time.sleep(1)
                log_message("✅ Modal fechado", "INFO")
        except NoSuchElementException:
            log_message("ℹ️ Nenhum modal detectado", "INFO")

    def acessar_baixa_pagamentos_lote(self):
        """Acessa a página de baixa de pagamentos em lote"""
        try:
            log_message("Acessando 'Baixa de pagamentos em lote'...", "INFO")
            link_baixa = self.wait.until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//a[contains(@class, 'setupAjax') and contains(text(), 'Baixa de pagamentos em lote')]"
                ))
            )
            link_baixa.click()
            time.sleep(2)
            log_message("✅ Página de baixa de pagamentos acessada", "SUCCESS")
        except Exception as e:
            log_message(f"❌ Erro ao acessar baixa de pagamentos: {e}", "ERROR")
            raise

    def processar_lote(self, numero_lote, cancel_flag=None):
        """Processa um lote específico"""
        try:
            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada pelo usuário", "WARNING")
                return False

            log_message(f"➡️ Processando lote: {numero_lote}", "INFO")

            # Localizar campo de número do lote
            log_message("🔍 Localizando campo de número do lote...", "INFO")
            campo_lote = self.wait.until(EC.presence_of_element_located((By.ID, "numeroLote")))
            campo_lote.clear()
            campo_lote.send_keys(numero_lote)
            log_message(f"⌨️ Número do lote '{numero_lote}' inserido", "INFO")
            time.sleep(0.5)

            # Clicar no botão de pesquisa
            log_message("🔎 Clicando no botão de pesquisa...", "INFO")
            try:
                botao_pesquisa = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "pesquisarBaixaLote"))
                )
                botao_pesquisa.click()
                log_message("✅ Botão de pesquisa clicado", "INFO")
            except Exception as e:
                log_message(f"⚠️ Erro ao clicar. Tentando JavaScript: {e}", "WARNING")
                botao_pesquisa = self.driver.find_element(By.ID, "pesquisarBaixaLote")
                self.driver.execute_script("arguments[0].click();", botao_pesquisa)
                log_message("✅ Botão clicado via JavaScript", "INFO")

            # Aguardar carregamento
            time.sleep(2)

            # Verificar resultados
            log_message("📋 Verificando resultados...", "INFO")
            try:
                tabela = self.driver.find_element(By.ID, "tabelaBaixaLote")
                linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody tr")

                if not linhas or len(linhas) == 0:
                    log_message(f"⚠️ Nenhum resultado encontrado para o lote {numero_lote}", "WARNING")
                    return False

                log_message(f"📊 Encontradas {len(linhas)} linha(s) para o lote {numero_lote}", "INFO")

                # Marcar checkbox "Selecionar Todos"
                log_message("☑️ Marcando checkbox 'Selecionar Todos'...", "INFO")
                try:
                    checkbox_todos = self.wait.until(
                        EC.element_to_be_clickable((By.ID, "selecionarTodos"))
                    )
                    if not checkbox_todos.is_selected():
                        checkbox_todos.click()
                        log_message("✅ Checkbox marcado", "INFO")
                    else:
                        log_message("ℹ️ Checkbox já estava marcado", "INFO")
                except Exception as e:
                    log_message(f"⚠️ Erro ao marcar checkbox: {e}", "WARNING")
                    checkbox_todos = self.driver.find_element(By.ID, "selecionarTodos")
                    self.driver.execute_script("arguments[0].click();", checkbox_todos)

                time.sleep(1)

                # Clicar no botão "Baixar Selecionados"
                log_message("💾 Clicando no botão 'Baixar Selecionados'...", "INFO")
                try:
                    botao_baixar = self.wait.until(
                        EC.element_to_be_clickable((By.ID, "baixarSelecionados"))
                    )
                    botao_baixar.click()
                    log_message("✅ Botão 'Baixar Selecionados' clicado", "INFO")
                except Exception as e:
                    log_message(f"⚠️ Erro ao clicar. Tentando JavaScript: {e}", "WARNING")
                    botao_baixar = self.driver.find_element(By.ID, "baixarSelecionados")
                    self.driver.execute_script("arguments[0].click();", botao_baixar)

                time.sleep(2)

                # Aguardar confirmação
                try:
                    self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "swal2-confirm")))
                    botao_confirmar = self.driver.find_element(By.CLASS_NAME, "swal2-confirm")
                    botao_confirmar.click()
                    log_message("✅ Confirmação realizada", "SUCCESS")
                    time.sleep(1)
                except TimeoutException:
                    log_message("ℹ️ Nenhuma confirmação necessária", "INFO")

                log_message(f"✅ Lote {numero_lote} processado com sucesso", "SUCCESS")
                return True

            except NoSuchElementException:
                log_message(f"⚠️ Tabela não encontrada para o lote {numero_lote}", "WARNING")
                return False

        except Exception as e:
            log_message(f"❌ Erro ao processar lote {numero_lote}: {e}", "ERROR")
            return False

    def processar_multiplos_lotes(self, lotes, cancel_flag=None):
        """Processa múltiplos lotes"""
        resultados = {
            "sucesso": [],
            "falha": [],
            "sem_resultados": []
        }

        total = len(lotes)
        log_message(f"📦 Iniciando processamento de {total} lote(s)", "INFO")

        for idx, lote in enumerate(lotes, start=1):
            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada pelo usuário", "WARNING")
                break

            log_message(f"\n{'=' * 60}", "INFO")
            log_message(f"📦 PROCESSANDO LOTE {idx} de {total}: {lote}", "INFO")
            log_message(f"{'=' * 60}\n", "INFO")

            sucesso = self.processar_lote(lote, cancel_flag)

            if sucesso:
                resultados["sucesso"].append(lote)
            else:
                resultados["sem_resultados"].append(lote)

            # Aguardar antes do próximo lote
            if idx < total:
                time.sleep(1)
                # Retornar à página de baixa de pagamentos
                try:
                    self.acessar_baixa_pagamentos_lote()
                except Exception as e:
                    log_message(f"⚠️ Erro ao retornar à página inicial: {e}", "WARNING")

        return resultados

    def exibir_resumo(self, resultados):
        """Exibe resumo do processamento"""
        log_message("\n" + "=" * 60, "INFO")
        log_message("📊 RESUMO FINAL DO PROCESSAMENTO", "INFO")
        log_message("=" * 60, "INFO")
        log_message(f"✅ Lotes processados com sucesso: {len(resultados['sucesso'])}", "SUCCESS")
        log_message(f"⚠️ Lotes sem resultados: {len(resultados['sem_resultados'])}", "WARNING")
        log_message(f"❌ Lotes com falha: {len(resultados['falha'])}", "ERROR")
        log_message(f"📦 Total processado: {sum(len(v) for v in resultados.values())}", "INFO")

        if resultados['sucesso']:
            log_message("\n✅ Lotes processados:", "SUCCESS")
            for lote in resultados['sucesso']:
                log_message(f"  - {lote}", "SUCCESS")

        if resultados['sem_resultados']:
            log_message("\n⚠️ Lotes sem resultados:", "WARNING")
            for lote in resultados['sem_resultados']:
                log_message(f"  - {lote}", "WARNING")

        if resultados['falha']:
            log_message("\n❌ Lotes com falha:", "ERROR")
            for lote in resultados['falha']:
                log_message(f"  - {lote}", "ERROR")

    def fechar_navegador(self):
        """Fecha o navegador"""
        if self.driver:
            log_message("Fechando navegador...", "INFO")
            self.driver.quit()
            log_message("✅ Navegador fechado", "SUCCESS")

    def run(self, params: dict):
        """Executa o processo completo de baixa de lotes"""
        username = params.get("username")
        password = params.get("password")
        lotes = params.get("lotes", [])  # Lista de números de lote
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode", False)

        if not lotes:
            log_message("❌ Nenhum lote fornecido para processamento", "ERROR")
            return

        try:
            log_message("🚀 Iniciando automação de baixa de lotes...", "INFO")

            # Inicializar driver
            self.inicializar_driver(headless_mode)

            # Fazer login
            self.fazer_login(username, password)

            # Navegar para módulo financeiro
            self.navegar_para_modulo_financeiro()

            # Fechar modal se necessário
            self.fechar_modal_se_necessario()

            # Acessar baixa de pagamentos
            self.acessar_baixa_pagamentos_lote()

            # Processar lotes
            resultados = self.processar_multiplos_lotes(lotes, cancel_flag)

            # Exibir resumo
            self.exibir_resumo(resultados)

            log_message("✅ Automação concluída com sucesso", "SUCCESS")

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            raise
        finally:
            self.fechar_navegador()


def run(params: dict):
    """Função de entrada para execução do módulo"""
    module = BaixaLoteModule()
    module.run(params)