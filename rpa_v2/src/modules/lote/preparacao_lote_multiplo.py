import os
import pandas as pd
import time
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from dotenv import load_dotenv

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule
from src.modules.lote.envio_lote_unimed import XMLGeneratorAutomation

load_dotenv()


class PreparacaoLoteMultiploModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Preparação de Lote Múltiplo")
        self.max_exames_por_lote = 99

    def dividir_exames_em_lotes(self, exames: list) -> list:
        """Divide a lista de exames em lotes de até 99 exames"""
        lotes = []
        for i in range(0, len(exames), self.max_exames_por_lote):
            lote = exames[i:i + self.max_exames_por_lote]
            lotes.append(lote)
        return lotes

    def get_unique_exames(self, file_path: str, modo_busca: str) -> list:
        if modo_busca == "exame":
            df = pd.read_excel(file_path)
            unique_exames = df['Exame'].dropna().unique().tolist()
        elif modo_busca == "guia":
            df = pd.read_excel(file_path)
            unique_exames = df['N Guia'].dropna().unique().tolist()
        else:
            raise ValueError("Modo de busca inválido. Use 'exame' ou 'guia'.")
        return unique_exames

    def voltar_tela_inicial_preparacao(self, driver, wait):
        """Retorna à tela inicial de preparação de exames"""
        try:
            log_message("🔄 Retornando à tela inicial de preparação...", "INFO")
            wait.until(EC.element_to_be_clickable((
                By.XPATH, "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]"
            ))).click()
            time.sleep(1)
            log_message("✅ Retornou à tela de preparação", "SUCCESS")
        except Exception as e:
            log_message(f"⚠️ Erro ao retornar à tela inicial: {e}", "WARNING")

    def processar_lote(self, driver, wait, exames_lote: list, modo_busca: str, cancel_flag):
        """Processa um lote de até 100 exames"""
        resultados_lote = []

        for exame in exames_lote:
            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada pelo usuário.", "WARNING")
                break

            try:
                log_message(f"➡️ Processando {modo_busca}: {exame}", "INFO")
                campo_id = "numeroGuia" if modo_busca == "guia" else "numeroExame"

                campo_exame = wait.until(EC.presence_of_element_located((By.ID, campo_id)))
                campo_exame.clear()
                campo_exame.send_keys(exame)
                time.sleep(0.5)

                try:
                    botao_pesquisa = wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento")))
                    try:
                        botao_pesquisa.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", botao_pesquisa)
                except Exception as e:
                    time.sleep(1)
                    botao_retry = driver.find_element(By.ID, "pesquisaFaturamento")
                    driver.execute_script("arguments[0].removeAttribute('disabled');", botao_retry)
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_retry)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", botao_retry)

                try:
                    modal_carregando = driver.find_element(By.XPATH,
                                                           "//div[contains(@class,'modal-body') and contains(., 'Carregando')]")
                    if modal_carregando.is_displayed():
                        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                except Exception:
                    pass

                time.sleep(1)

                tbody_rows = driver.find_elements(By.CSS_SELECTOR, "#tabelaPreFaturamentoTbody tr")
                if len(tbody_rows) == 0:
                    log_message(f"⚠️ Nenhum resultado encontrado para {exame}. Pulando.", "WARNING")
                    resultados_lote.append({"exame": exame, "status": "sem_resultados"})
                    continue

                time.sleep(1)

                try:
                    checkbox = wait.until(EC.element_to_be_clickable((By.ID, "checkTodosPreFaturar")))
                    try:
                        checkbox.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", checkbox)
                except Exception:
                    time.sleep(1)
                    checkbox_retry = driver.find_element(By.ID, "checkTodosPreFaturar")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox_retry)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", checkbox_retry)

                try:
                    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                except Exception:
                    pass

                time.sleep(1)

                try:
                    acoes_btn = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//a[contains(@class, 'toggleMaisDeUm') and contains(., 'Ações')]"
                    )))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", acoes_btn)
                    time.sleep(0.5)
                    try:
                        acoes_btn.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", acoes_btn)
                except Exception:
                    time.sleep(1)
                    acoes_retry = driver.find_element(By.XPATH,
                                                      "//a[contains(@class, 'toggleMaisDeUm') and contains(., 'Ações')]")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", acoes_retry)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", acoes_retry)

                time.sleep(1)

                driver.execute_script("""
                    const onlineBtn = document.querySelector("a[data-url*='statusConferido=O']");
                    if (onlineBtn) { onlineBtn.click(); }
                """)

                time.sleep(1)

                if modo_busca == "guia":
                    try:
                        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                    except Exception:
                        time.sleep(1)

                resultados_lote.append({"exame": exame, "status": "sucesso"})
                log_message(f"✅ {modo_busca.title()} {exame} processado com sucesso.", "SUCCESS")

            except Exception as e:
                resultados_lote.append({"exame": exame, "status": "erro", "erro": str(e)})
                log_message(f"❌ Erro ao processar {exame}: {e}", "ERROR")

        return resultados_lote

    def gerar_ou_enviar_lote(self, driver, wait, gera_xml_tiss: str, username: str, password: str,
                             unimed_user: str, unimed_pass: str, pasta_download: str,
                             headless_mode: bool, cancel_flag, modo_busca: str, numero_lote: int):
        """Gera ou envia o lote após preparação"""
        if gera_xml_tiss == "sim":
            log_message(f"📤 Gerando e enviando XML para Unimed - Lote {numero_lote}...", "INFO")
            automacao = XMLGeneratorAutomation(username, password, pasta_download=pasta_download,headless=headless_mode)
            sucesso_envio = automacao.executar_processo_completo_login_navegacao(unimed_user, unimed_pass, cancel_flag=cancel_flag)
            if not sucesso_envio:
                log_message(f"❌ Falha ao processar/enviar lote {numero_lote} para Unimed.", "ERROR")
                return False
            else:
                log_message(f"✅ Lote {numero_lote} enviado para Unimed com sucesso!", "SUCCESS")
                return True

        elif gera_xml_tiss == "nao":
            log_message(f"📋 Preparando lote {numero_lote} para geração manual...", "INFO")
            try:
                campo_id = "numeroGuia" if modo_busca == "guia" else "numeroExame"
                campo_exame = wait.until(EC.presence_of_element_located((By.ID, campo_id)))
                campo_exame.clear()
                time.sleep(0.5)

                select2_container = wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, ".select2-selection[aria-labelledby*='convenioId']"))
                )
                select2_container.click()
                time.sleep(1)

                opcao_unimed = wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH,
                         "//li[contains(@class, 'select2-results__option') and text()='UNIMED (LONDRINA)']"))
                )
                opcao_unimed.click()
                time.sleep(1)

                select_element = wait.until(EC.presence_of_element_located((By.ID, "conferido")))
                select_conferido = Select(select_element)
                select_conferido.select_by_value("O")
                time.sleep(1)

                botao_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento")))
                botao_pesquisar.click()
                time.sleep(2)

                tempo_maximo = time.time() + 60
                while time.time() < tempo_maximo:
                    try:
                        modal_carregando = driver.find_element(By.XPATH,
                                                               "//div[contains(@class,'modal-body') and contains(., 'Carregando')]")
                        if modal_carregando.is_displayed():
                            time.sleep(1)
                        else:
                            break
                    except Exception:
                        break

                try:
                    gerar_tiss_checkbox = driver.find_element(By.ID, "gerarArquivoTiss")
                    if gerar_tiss_checkbox.is_selected():
                        gerar_tiss_checkbox.click()
                        log_message("Checkbox 'gerarArquivoTiss' desmarcado.", "INFO")
                    time.sleep(1)
                except Exception as e:
                    log_message(f"Não foi possível desmarcar gerarArquivoTiss: {e}", "WARNING")

                try:
                    botao_situacao = wait.until(EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "a.btn.btn-danger[onclick*='modalFaturamento']")))
                    botao_situacao.click()
                    log_message(f"✅ Lote {numero_lote} preparado para geração manual.", "SUCCESS")
                    time.sleep(2)
                    return True
                except Exception as e:
                    log_message(f"❌ Erro ao preparar lote {numero_lote}: {e}", "ERROR")
                    return False

            except Exception as e:
                log_message(f"❌ Erro na rotina de geração manual - Lote {numero_lote}: {e}", "ERROR")
                return False

    def run(self, params: dict):
        username = params.get("username")
        password = params.get("password")
        excel_file = params.get("excel_file")
        modo_busca = params.get("modo_busca", "exame")
        cancel_flag = params.get("cancel_flag")
        gera_xml_tiss = params.get("gera_xml_tiss", "sim")
        headless_mode = params.get("headless_mode")
        unimed_user = params.get("unimed_user")
        unimed_pass = params.get("unimed_pass")
        pasta_download = params.get("pasta_download", os.path.join(os.getcwd(), "xml"))

        try:
            exames_unicos = self.get_unique_exames(excel_file, modo_busca)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o Excel: {e}")
            return

        if not exames_unicos:
            messagebox.showerror("Erro", "Nenhum exame encontrado no arquivo.")
            return

        # Dividir exames em lotes de 100
        lotes = self.dividir_exames_em_lotes(exames_unicos)
        total_lotes = len(lotes)

        log_message(f"📊 Total de exames únicos: {len(exames_unicos)}", "INFO")
        log_message(f"📦 Divididos em {total_lotes} lote(s) de até {self.max_exames_por_lote} exames", "INFO")

        url = os.getenv("SYSTEM_URL", "https://pathoweb.com.br/login/auth")
        driver = BrowserFactory.create_chrome(headless=headless_mode)
        wait = WebDriverWait(driver, 15)

        todos_resultados = []

        try:
            log_message("Iniciando automação de preparação de exames em múltiplos lotes...", "INFO")
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

            elif "moduloFaturamento" in current_url:
                log_message("✅ Já está no módulo de faturamento - pulando navegação", "SUCCESS")
            else:
                log_message(f"⚠️ URL inesperada detectada: {current_url}", "WARNING")
                driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                time.sleep(2)

            try:
                modal_close_button = driver.find_element(By.CSS_SELECTOR,
                                                         "#mensagemParaClienteModal .modal-footer button")
                if modal_close_button.is_displayed():
                    modal_close_button.click()
                    time.sleep(1)
            except Exception:
                pass

            wait.until(EC.element_to_be_clickable((
                By.XPATH, "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]"
            ))).click()
            time.sleep(1)

            # Processar cada lote sequencialmente
            for idx, lote in enumerate(lotes, start=1):
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break

                log_message(f"\n{'=' * 60}", "INFO")
                log_message(f"📦 PROCESSANDO LOTE {idx} de {total_lotes} ({len(lote)} exames)", "INFO")
                log_message(f"{'=' * 60}\n", "INFO")

                # Processar o lote atual
                resultados_lote = self.processar_lote(driver, wait, lote, modo_busca, cancel_flag)
                todos_resultados.extend(resultados_lote)

                # Resumo do lote atual
                sucesso_lote = [r for r in resultados_lote if r["status"] == "sucesso"]
                erro_lote = [r for r in resultados_lote if r["status"] == "erro"]
                sem_resultados_lote = [r for r in resultados_lote if r["status"] == "sem_resultados"]

                log_message(f"\n📊 Resumo do Lote {idx}:", "INFO")
                log_message(f"✅ Sucesso: {len(sucesso_lote)}", "SUCCESS")
                log_message(f"⚠️ Sem resultados: {len(sem_resultados_lote)}", "WARNING")
                log_message(f"❌ Erros: {len(erro_lote)}", "ERROR")

                # Gerar/enviar lote após processamento
                self.gerar_ou_enviar_lote(driver, wait, gera_xml_tiss, username, password,
                                          unimed_user, unimed_pass, pasta_download, headless_mode,
                                          cancel_flag, modo_busca, idx)

                # Se não for o último lote, retornar à tela inicial
                if idx < total_lotes:
                    self.voltar_tela_inicial_preparacao(driver, wait)

            # Resumo final consolidado
            total = len(todos_resultados)
            sucesso_total = [r for r in todos_resultados if r["status"] == "sucesso"]
            erro_total = [r for r in todos_resultados if r["status"] == "erro"]
            sem_resultados_total = [r for r in todos_resultados if r["status"] == "sem_resultados"]

            log_message("\n" + "=" * 60, "INFO")
            log_message("📊 RESUMO FINAL - TODOS OS LOTES", "INFO")
            log_message("=" * 60, "INFO")
            log_message(f"Total de exames processados: {total}", "INFO")
            log_message(f"✅ Sucesso: {len(sucesso_total)}", "SUCCESS")
            log_message(f"⚠️ Sem resultados: {len(sem_resultados_total)}", "WARNING")
            log_message(f"❌ Erros: {len(erro_total)}", "ERROR")
            log_message(f"📦 Total de lotes gerados: {total_lotes}", "INFO")

            messagebox.showinfo("Sucesso",
                                f"✅ Processamento finalizado!\n\n"
                                f"📦 Lotes processados: {total_lotes}\n"
                                f"📊 Total de exames: {total}\n"
                                f"✅ Sucesso: {len(sucesso_total)}\n"
                                f"⚠️ Sem resultados: {len(sem_resultados_total)}\n"
                                f"❌ Erros: {len(erro_total)}"
                                )

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            driver.quit()


def run(params: dict):
    module = PreparacaoLoteMultiploModule()
    module.run(params)