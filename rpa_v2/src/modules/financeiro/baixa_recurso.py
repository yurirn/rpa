import os
import pandas as pd
import time
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from dotenv import load_dotenv

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

load_dotenv()


class BaixaRecursoModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Baixa de Recurso")

    def get_unique_exames(self, file_path: str) -> list:
        """Extrai lista de exames únicos do arquivo Excel"""
        df = pd.read_excel(file_path)
        unique_exames = df['Exame'].dropna().unique().tolist()
        return unique_exames

    def fechar_modal_se_necessario(self, driver):
        """Fecha modal de mensagem se estiver visível"""
        try:
            modal_close_button = driver.find_element(
                By.CSS_SELECTOR,
                "#mensagemParaClienteModal .modal-footer button"
            )
            if modal_close_button.is_displayed():
                modal_close_button.click()
                time.sleep(1)
                log_message("✅ Modal fechado com sucesso", "INFO")
        except Exception:
            log_message("ℹ️ Nenhum modal para fechar", "INFO")

    def navegar_para_modulo_faturamento(self, driver, wait):
        """Navega para o módulo de faturamento se necessário"""
        log_message("Verificando se precisa navegar para módulo de faturamento...", "INFO")
        current_url = driver.current_url

        if current_url == "https://dap.pathoweb.com.br/" or "trocarModulo" in current_url:
            log_message("Detectada tela de seleção de módulos - navegando para módulo de faturamento...", "INFO")
            try:
                modulo_link = wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=2']")
                    )
                )
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

    def fazer_login(self, driver, wait, username, password):
        """Realiza login no sistema"""
        log_message("Iniciando processo de login...", "INFO")

        url = os.getenv("SYSTEM_URL", "https://dap.pathoweb.com.br/login/auth")
        driver.get(url)

        wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

        log_message("✅ Login realizado com sucesso", "SUCCESS")
        time.sleep(2)

    def acessar_tela_recursos(self, driver, wait):
        """Acessa a tela de recursos"""
        log_message("Acessando tela de recursos...", "INFO")

        try:
            link_recursos = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[contains(@class, 'setupAjax') and contains(text(), 'Recursos')]")
                )
            )
            link_recursos.click()
            time.sleep(1)
            log_message("✅ Tela de recursos acessada com sucesso", "SUCCESS")
        except Exception as e:
            log_message(f"❌ Erro ao acessar tela de recursos: {e}", "ERROR")
            raise

    def limpar_campo_busca(self, driver, wait):
        """Limpa o campo de busca de exame"""
        try:
            campo_exame = wait.until(EC.presence_of_element_located((By.ID, "numeroExame")))
            campo_exame.clear()
            log_message("🧹 Campo de busca limpo", "INFO")
            time.sleep(0.3)
        except Exception as e:
            log_message(f"⚠️ Erro ao limpar campo de busca: {e}", "WARNING")

    def pesquisar_exame(self, driver, wait, numero_exame):
        """Pesquisa um exame específico"""
        log_message(f"🔍 Pesquisando exame: {numero_exame}", "INFO")

        try:
            # Localizar e preencher campo
            campo_exame = wait.until(EC.presence_of_element_located((By.ID, "numeroExame")))
            campo_exame.clear()
            campo_exame.send_keys(numero_exame)
            time.sleep(0.5)

            # Clicar no botão de pesquisa
            botao_pesquisa = wait.until(
                EC.element_to_be_clickable((By.ID, "pesquisaRecurso"))
            )

            try:
                botao_pesquisa.click()
                log_message("✅ Botão de pesquisa clicado", "INFO")
            except Exception:
                driver.execute_script("arguments[0].click();", botao_pesquisa)
                log_message("✅ Botão de pesquisa clicado (JavaScript)", "INFO")

            time.sleep(2)

        except Exception as e:
            log_message(f"❌ Erro ao pesquisar exame: {e}", "ERROR")
            raise

    def aguardar_modal_carregamento(self, driver):
        """Aguarda o fechamento do modal de carregamento"""
        try:
            modal_carregando = driver.find_element(
                By.XPATH,
                "//div[contains(@class,'modal-body') and contains(., 'Carregando')]"
            )
            if modal_carregando.is_displayed():
                log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                WebDriverWait(driver, 30).until(
                    EC.invisibility_of_element_located((By.ID, "spinner"))
                )
                log_message("✅ Modal de carregamento fechado", "INFO")
        except Exception:
            log_message("ℹ️ Modal não detectado", "INFO")

    def validar_resultados_tabela(self, driver, numero_exame):
        """Valida se há resultados na tabela"""
        log_message("📋 Validando resultados da tabela...", "INFO")

        try:
            tbody_rows = driver.find_elements(By.CSS_SELECTOR, "#tabelaRecurso tbody tr")
            log_message(f"📊 Encontradas {len(tbody_rows)} linha(s) na tabela", "INFO")

            if len(tbody_rows) == 0:
                log_message(f"⚠️ Nenhum resultado encontrado para exame {numero_exame}", "WARNING")
                return False

            return True

        except Exception as e:
            log_message(f"⚠️ Erro ao validar resultados: {e}", "WARNING")
            return False

    def marcar_checkbox_todos(self, driver, wait):
        """Marca o checkbox para selecionar todos os recursos"""
        log_message("☑️ Marcando checkbox 'Selecionar Todos'...", "INFO")

        max_tentativas = 3
        for tentativa in range(1, max_tentativas + 1):
            try:
                checkbox = wait.until(
                    EC.element_to_be_clickable((By.ID, "checkTodosRecursos"))
                )

                try:
                    checkbox.click()
                    log_message("✅ Checkbox marcado (click normal)", "INFO")
                    return True
                except Exception:
                    driver.execute_script("arguments[0].click();", checkbox)
                    log_message("✅ Checkbox marcado (JavaScript)", "INFO")
                    return True

            except Exception as e:
                log_message(f"⚠️ Tentativa {tentativa} falhou: {e}", "WARNING")
                if tentativa < max_tentativas:
                    time.sleep(1)
                else:
                    raise

        return False

    def clicar_botao_acoes(self, driver, wait):
        """Clica no botão de ações"""
        log_message("🎬 Clicando no botão 'Ações'...", "INFO")

        try:
            acoes_btn = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(@class, 'btn-primary') and contains(., 'Ações')]")
                )
            )

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", acoes_btn)
            time.sleep(0.5)

            try:
                acoes_btn.click()
                log_message("✅ Botão 'Ações' clicado (click normal)", "INFO")
            except Exception:
                driver.execute_script("arguments[0].click();", acoes_btn)
                log_message("✅ Botão 'Ações' clicado (JavaScript)", "INFO")

            time.sleep(1)

        except Exception as e:
            log_message(f"❌ Erro ao clicar no botão 'Ações': {e}", "ERROR")
            raise

    def selecionar_baixar_recursos(self, driver):
        """Seleciona a opção 'Baixar recursos'"""
        log_message("📥 Selecionando 'Baixar recursos'...", "INFO")

        try:
            driver.execute_script("""
                const baixarBtn = document.querySelector("a[onclick*='baixarRecursos']");
                if (baixarBtn) { 
                    baixarBtn.click(); 
                }
            """)
            log_message("✅ Opção 'Baixar recursos' selecionada", "SUCCESS")
            time.sleep(1)

        except Exception as e:
            log_message(f"❌ Erro ao selecionar 'Baixar recursos': {e}", "ERROR")
            raise

    def processar_exame(self, driver, wait, numero_exame):
        """Processa um exame completo"""
        try:
            log_message(f"➡️ Processando exame: {numero_exame}", "INFO")

            self.limpar_campo_busca(driver, wait)
            self.pesquisar_exame(driver, wait, numero_exame)
            self.aguardar_modal_carregamento(driver)

            if not self.validar_resultados_tabela(driver, numero_exame):
                return {"exame": numero_exame, "status": "sem_resultados"}

            time.sleep(1)

            self.marcar_checkbox_todos(driver, wait)

            self.aguardar_modal_carregamento(driver)
            time.sleep(1)

            self.clicar_botao_acoes(driver, wait)
            self.selecionar_baixar_recursos(driver)

            self.aguardar_modal_carregamento(driver)
            time.sleep(1)

            log_message(f"✅ Exame {numero_exame} processado com sucesso", "SUCCESS")
            return {"exame": numero_exame, "status": "sucesso"}

        except Exception as e:
            log_message(f"❌ Erro ao processar exame {numero_exame}: {e}", "ERROR")
            return {"exame": numero_exame, "status": "erro", "erro": str(e)}

    def exibir_resumo_final(self, resultados):
        """Exibe resumo final do processamento"""
        total = len(resultados)
        sucesso = [r for r in resultados if r["status"] == "sucesso"]
        erro = [r for r in resultados if r["status"] == "erro"]
        sem_resultados = [r for r in resultados if r["status"] == "sem_resultados"]
        erro_validacao = [r for r in resultados if r["status"] == "erro_validacao"]

        log_message("\n" + "=" * 60, "INFO")
        log_message("📊 RESUMO FINAL DO PROCESSAMENTO", "INFO")
        log_message("=" * 60, "INFO")
        log_message(f"Total de exames: {total}", "INFO")
        log_message(f"✅ Sucesso: {len(sucesso)}", "SUCCESS")
        log_message(f"⚠️ Sem resultados: {len(sem_resultados)}", "WARNING")
        log_message(f"⚠️ Erro validação: {len(erro_validacao)}", "WARNING")
        log_message(f"❌ Erros: {len(erro)}", "ERROR")

        messagebox.showinfo(
            "Sucesso",
            f"✅ Processamento finalizado!\n\n"
            f"Total: {total}\n"
            f"Sucesso: {len(sucesso)}\n"
            f"Sem resultados: {len(sem_resultados)}\n"
            f"Erros: {len(erro) + len(erro_validacao)}"
        )

    def run(self, params: dict):
        """Executa o processo completo de baixa de recursos"""
        username = params.get("username")
        password = params.get("password")
        excel_file = params.get("excel_file")
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode", False)

        # Validar parâmetros obrigatórios
        if not all([username, password, excel_file]):
            messagebox.showerror("Erro", "Parâmetros obrigatórios não fornecidos")
            return

        try:
            exames_unicos = self.get_unique_exames(excel_file)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o Excel: {e}")
            return

        if not exames_unicos:
            messagebox.showerror("Erro", "Nenhum exame encontrado no arquivo.")
            return

        log_message(f"📊 Total de exames únicos encontrados: {len(exames_unicos)}", "INFO")

        driver = BrowserFactory.create_chrome(headless=headless_mode)
        wait = WebDriverWait(driver, 15)
        resultados = []

        try:
            log_message("🚀 Iniciando automação de baixa de recursos...", "INFO")

            self.fazer_login(driver, wait, username, password)
            self.navegar_para_modulo_faturamento(driver, wait)
            self.fechar_modal_se_necessario(driver)
            self.acessar_tela_recursos(driver, wait)

            # Processar cada exame
            for exame in exames_unicos:
                if cancel_flag and cancel_flag.is_set():
                    log_message("⚠️ Execução cancelada pelo usuário", "WARNING")
                    break

                resultado = self.processar_exame(driver, wait, exame)
                resultados.append(resultado)

            self.exibir_resumo_final(resultados)

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            log_message("🔒 Fechando navegador...", "INFO")
            driver.quit()


def run(params: dict):
    """Função de entrada para execução do módulo"""
    module = BaixaRecursoModule()
    module.run(params)