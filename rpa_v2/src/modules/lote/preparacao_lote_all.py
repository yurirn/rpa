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

load_dotenv()

class PreparacaoLoteModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Preparação de Lote")

    def get_unique_exames(self, file_path: str, modo_busca: str) -> dict:
        if modo_busca == "exame":
            df = pd.read_excel(file_path)
            df = df.dropna(subset=['Exame'])

            # Pegar primeiro registro de cada exame único (mais eficiente)
            df_unique = df.drop_duplicates(subset=['Exame'], keep='first')

            exames_info = {}
            for _, row in df_unique.iterrows():
                exames_info[row['Exame']] = {
                    'convenio': row['Convenio'],
                    'procedencia': row['Procedência']
                }

            log_message(f"Encontrados {len(exames_info)} exames únicos", "INFO")
            return exames_info

        elif modo_busca == "guia":
            df = pd.read_excel(file_path)
            df = df.dropna(subset=['N Guia'])  # Usar coluna da guia
            guias_info = {}
            for _, row in df.iterrows():
                guia = str(row['N Guia'])  # Converter para string
                if guia not in guias_info:
                    guias_info[guia] = {
                        'convenio': row['Convenio'],
                        'procedencia': row['Procedência']
                    }
            log_message(f"Encontradas {len(guias_info)} guias únicas", "INFO")
            return guias_info
        else:
            raise ValueError("Modo de busca inválido. Use 'exame' ou 'guia'.")

    def perform_login(self, driver, wait, username, password, url):
        """
        Realiza o login no sistema
        """
        log_message("Iniciando automação de preparação de exames...", "INFO")
        driver.get(url)
        wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

    def validate_and_navigate_module(self, driver, wait):
        """
        Verifica e navega para o módulo de faturamento se necessário
        """
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
                # Tentar navegar diretamente pela URL como fallback
                driver.get("https://dap.pathoweb.com.br/moduloFaturamento/index")
                time.sleep(2)
                log_message("🔄 Navegação direta para módulo realizada", "INFO")

        elif "moduloFaturamento" in current_url:
            log_message("✅ Já está no módulo de faturamento - pulando navegação", "SUCCESS")
        else:
            log_message(f"⚠️ URL inesperada detectada: {current_url}", "WARNING")
            # Tentar navegar diretamente como fallback
            driver.get("https://dap.pathoweb.com.br/moduloFaturamento/index")
            time.sleep(2)
            log_message("🔄 Navegação direta para módulo realizada (fallback)", "INFO")

    def navigate_to_exam_preparation(self, driver, wait):
        """
        Navega para a funcionalidade de preparar exames para fatura
        """
        try:
            modal_close_button = driver.find_element(By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button")
            if modal_close_button.is_displayed():
                modal_close_button.click()
                time.sleep(1)
        except Exception:
            pass

        wait.until(EC.element_to_be_clickable((
            By.XPATH, "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]"
        ))).click()
        time.sleep(1)

    def fechar_modais_interferentes(self, driver):
        """
        Fecha/remover modais/backdrops que bloqueiam cliques.
        """
        try:
            driver.execute_script("""
                const modais = document.querySelectorAll('.modal.in, .modal.show');
                modais.forEach(m => {
                    m.style.display = 'none';
                    m.classList.remove('in', 'show');
                });
                document.querySelectorAll('.modal-backdrop').forEach(b => b.remove());
            """)
            time.sleep(0.5)
            log_message("Modais/backdrops fechados para liberar a tela.", "INFO")
        except Exception as e:
            log_message(f"Não foi possível fechar modais: {e}", "WARNING")

    def process_single_exam(self, driver, wait, exame, modo_busca):
        """
        Processa um único exame na preparação
        """
        try:
            log_message(f"➡️ Processando {modo_busca}: {exame}", "INFO")

            select_element2 = wait.until(EC.presence_of_element_located((By.ID, "cobrarDe")))
            select_cobrar_de = Select(select_element2)
            select_cobrar_de.select_by_value("")
            time.sleep(0.5)

            campo_id = "numeroGuia" if modo_busca == "guia" else "numeroExame"
            campo_exame = wait.until(EC.presence_of_element_located((By.ID, campo_id)))
            campo_exame.clear()
            campo_exame.send_keys(exame)

            # Garantir que nenhum modal/backdrop esteja bloqueando o clique
            self.fechar_modais_interferentes(driver)
            wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento"))).click()

            # Aguardar modal de carregamento
            try:
                modal_carregando = driver.find_element(By.XPATH,
                                                       "//div[contains(@class,'modal-body') and contains(., 'Carregando')]")
                if modal_carregando.is_displayed():
                    log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                    log_message("✅ Modal de carregamento fechado", "INFO")
            except Exception:
                log_message("ℹ️ Modal não detectado. Prosseguindo...", "INFO")

            time.sleep(1.5)

            # Validar se há resultados
            try:
                tbody_rows = driver.find_elements(By.CSS_SELECTOR, "#tabelaPreFaturamentoTbody tr")
                if len(tbody_rows) == 0:
                    log_message(f"⚠️ Nenhum resultado encontrado para {exame}. Pulando.", "WARNING")
                    return {"exame": exame, "status": "sem_resultados"}
            except Exception as e:
                log_message(f"⚠️ Erro ao validar resultados da tabela: {e}", "WARNING")
                return {"exame": exame, "status": "erro_validacao", "erro": str(e)}

            # Selecionar todos e processar
            wait.until(EC.element_to_be_clickable((By.ID, "checkTodosPreFaturar"))).click()
            acoes_btn = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//a[contains(@class, 'toggleMaisDeUm') and contains(., 'Ações')]"
            )))
            acoes_btn.click()
            time.sleep(1)

            # Executar script para marcar como online
            driver.execute_script("""
                const onlineBtn = document.querySelector("a[data-url*='statusConferido=O']");
                if (onlineBtn) { onlineBtn.click(); }
            """)
            time.sleep(1)

            # Aguardar carregamento adicional para modo guia
            if modo_busca == "guia":
                try:
                    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                    log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                    log_message("✅ Modal de carregamento fechado", "INFO")
                except Exception:
                    log_message("ℹ️ Modal não detectado. Prosseguindo...", "INFO")
                    time.sleep(1)

            log_message(f"✅ {modo_busca.title()} {exame} processado com sucesso.", "SUCCESS")
            return {"exame": exame, "status": "sucesso"}

        except Exception as e:
            log_message(f"⌚ Erro ao processar {exame}: {e}", "ERROR")
            return {"exame": exame, "status": "erro", "erro": str(e)}

    def show_processing_results(self, resultados):
        """
        Mostra o resumo dos resultados do processamento
        """
        total = len(resultados)
        sucesso = [r for r in resultados if r["status"] == "sucesso"]
        erro = [r for r in resultados if r["status"] == "erro"]
        sem_resultados = [r for r in resultados if r["status"] == "sem_resultados"]
        erro_validacao = [r for r in resultados if r["status"] == "erro_validacao"]

        log_message("\nResumo do processamento:", "INFO")
        log_message(f"Total: {total}", "INFO")
        log_message(f"Sucesso: {len(sucesso)}", "SUCCESS")
        log_message(f"Sem resultados: {len(sem_resultados)}", "WARNING")
        log_message(f"Erro validação: {len(erro_validacao)}", "WARNING")
        log_message(f"Erro processamento: {len(erro)}", "ERROR")

        messagebox.showinfo("Sucesso",
                            f"✅ Processamento finalizado!\n"
                            f"Total: {total}\n"
                            f"Sucesso: {len(sucesso)}\n"
                            f"Sem resultados: {len(sem_resultados)}\n"
                            f"Erros: {len(erro) + len(erro_validacao)}"
                            )

    def preencher_filtros(self, driver, wait, modo_busca, cobrar_de, convenio, procedencia):
        """
        Preenche os filtros para geração de lote
        """
        log_message("Preenchendo filtros para geração de lote...", "INFO")

        # Limpar campo de pesquisa
        campo_id = "numeroGuia" if modo_busca == "guia" else "numeroExame"
        campo_exame = wait.until(EC.presence_of_element_located((By.ID, campo_id)))
        campo_exame.clear()
        time.sleep(0.5)

        select2_container_convenio = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, ".select2-selection[aria-labelledby*='convenioId']"))
        )
        select2_container_convenio.click()
        time.sleep(1)
        opcao_convenio = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//li[contains(@class, 'select2-results__option') and text()='{convenio}']"))
        )
        opcao_convenio.click()
        log_message(f"✅ Convenio {convenio} inserido com sucesso.", "SUCCESS")
        time.sleep(1)

        select2_container_procedencia = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, ".select2-selection[aria-labelledby*='procedenciaId']"))
        )
        select2_container_procedencia.click()
        time.sleep(1)
        opcao_procedencia = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//li[contains(@class, 'select2-results__option') and text()='{procedencia}']"))
        )
        opcao_procedencia.click()
        log_message(f"✅ Procedencia {procedencia} inserida com sucesso.", "SUCCESS")
        time.sleep(1)

        # Configurar seletor conferido
        select_element1 = wait.until(EC.presence_of_element_located((By.ID, "conferido")))
        select_conferido = Select(select_element1)
        select_conferido.select_by_value("O")
        log_message(f"✅ Filtro on-line inserido com sucesso.", "SUCCESS")
        time.sleep(1)

        # Configurar cobrar de - usando o parâmetro recebido
        select_element2 = wait.until(EC.presence_of_element_located((By.ID, "cobrarDe")))
        select_cobrar_de = Select(select_element2)
        select_cobrar_de.select_by_value(cobrar_de)  # C = Convênio, P = Procedência
        time.sleep(1)

        log_message(f"Filtros configurados - Cobrar de: {'Convênio' if cobrar_de == 'C' else 'Procedência'}", "INFO")

    def gerar_lote(self, driver, wait, modo_busca, cobrar_de, exames_info):
        """
        Executa a rotina de geração de lote manual usando dados da planilha
        """
        log_message("Executando rotina de geração de lote manual...", "INFO")

        primeiro_exame = next(iter(exames_info.keys()))
        convenio = exames_info[primeiro_exame]['convenio']
        procedencia = exames_info[primeiro_exame]['procedencia']

        try:
            # Preencher filtros com dados da planilha
            self.preencher_filtros(driver, wait, modo_busca, cobrar_de, convenio, procedencia)

            # Executar pesquisa
            botao_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento")))
            botao_pesquisar.click()
            time.sleep(2)

            # Aguardar finalização da pesquisa
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

            # Desmarcar checkbox gerarArquivoTiss
            try:
                gerar_tiss_checkbox = driver.find_element(By.ID, "gerarArquivoTiss")
                if gerar_tiss_checkbox.is_selected():
                    gerar_tiss_checkbox.click()
                    log_message("Checkbox 'gerarArquivoTiss' desmarcado.", "INFO")
                time.sleep(1)
            except Exception as e:
                log_message(f"Não foi possível desmarcar gerarArquivoTiss: {e}", "WARNING")

            gerar_lote = self.clicar_gerar_lote(driver, wait)
            return gerar_lote

        except Exception as e:
            log_message(f"Erro na rotina de geração de XML TISS manual: {e}", "ERROR")

    def clicar_gerar_lote(self, driver, wait):
        # Clicar no botão de situação de faturamento
        try:
            # 1. Aguardar o elemento estar PRESENTE no DOM
            log_message("Aguardando a presença do botão de download...", "INFO")
            botao = wait.until(
                EC.presence_of_element_located((By.ID, "executarMudancaSitFaturamento"))
            )

            driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", botao)
            time.sleep(1)

            # 3. Clicar no botão usando JavaScript
            log_message("Tentando clicar no botão via JavaScript...", "INFO")
            driver.execute_script("arguments[0].click();", botao)

            log_message("✅ Botão de download clicado com sucesso", "SUCCESS")

            modal_carregando = driver.find_element(By.XPATH,
                                                   "//div[contains(@class,'modal-body') and contains(., 'Carregando')]")
            if modal_carregando.is_displayed():
                time.sleep(1)

        except Exception as e:
            log_message(f"Não foi possível clicar no botão de situação de faturamento: {e}", "ERROR")
            return False

        return True

    def desmarcar_checkbox(self, driver):
        try:
            gerar_tiss_checkbox = driver.find_element(By.ID, "gerarArquivoTiss")
            if gerar_tiss_checkbox.is_selected():
                gerar_tiss_checkbox.click()
                log_message("Checkbox 'gerarArquivoTiss' desmarcado.", "INFO")
            time.sleep(1)
        except Exception as e:
            log_message(f"Não foi possível desmarcar gerarArquivoTiss: {e}", "WARNING")

    def run(self, params: dict):
        username = params.get("username")
        password = params.get("password")
        excel_file = params.get("excel_file")
        modo_busca = params.get("modo_busca", "exame")
        cancel_flag = params.get("cancel_flag")
        gera_xml_tiss = params.get("gera_xml_tiss", "sim")
        headless_mode = params.get("headless_mode")
        cobrar_de = params.get("cobrar_de", "C")

        try:
            exames_info = self.get_unique_exames(excel_file, modo_busca)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o Excel: {e}")
            return

        if not exames_info:
            messagebox.showerror("Erro", "Nenhum exame encontrado no arquivo.")
            return

        url = os.getenv("SYSTEM_URL", "https://dap.pathoweb.com.br/login/auth")
        driver = BrowserFactory.create_chrome(headless=headless_mode)
        wait = WebDriverWait(driver, 15)
        resultados = []

        try:
            # Realizar login
            self.perform_login(driver, wait, username, password, url)

            # Validar e navegar para módulo
            self.validate_and_navigate_module(driver, wait)

            # Navegar para preparar exames
            self.navigate_to_exam_preparation(driver, wait)

            for exame in exames_info.keys():
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break

                resultado = self.process_single_exam(driver, wait, exame, modo_busca)
                resultados.append(resultado)

            # Mostrar resultados
            self.show_processing_results(resultados)

            # Executar rotina de geração de lote se necessário
            sucesso_lote = self.gerar_lote(driver, wait, modo_busca, cobrar_de, exames_info)
            if sucesso_lote:
                messagebox.showinfo("Sucesso", "✅ Lote gerado com sucesso!")
                log_message("Lote gerado com sucesso!", "SUCCESS")
            else:
                messagebox.showerror("Erro", "❌ Erro ao gerar o lote. Verifique os logs.")
                log_message("Falha na geração do lote", "ERROR")

        except Exception as e:
            log_message(f"Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"Erro durante a automação:\n{e}")
        finally:
            driver.quit()

def run(params: dict):
    module = PreparacaoLoteModule()
    module.run(params)