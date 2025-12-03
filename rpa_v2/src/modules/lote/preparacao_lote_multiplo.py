import os
import pandas as pd
import time
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from dotenv import load_dotenv
from socket import timeout as TimeoutException

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

    def fechar_sweetalert(self, driver):
        """Fecha qualquer SweetAlert2 aberto que possa estar bloqueando a interação"""
        try:
            # Verificar se há SweetAlert visível
            sweetalert = driver.find_element(By.CLASS_NAME, "swal2-container")
            if sweetalert.is_displayed():
                log_message("⚠️ SweetAlert detectado - fechando...", "WARNING")

                # Tentar fechar pelo botão
                try:
                    botao_ok = driver.find_element(By.CSS_SELECTOR, ".swal2-confirm")
                    botao_ok.click()
                    time.sleep(0.5)
                    log_message("✅ SweetAlert fechado via botão", "INFO")
                except Exception:
                    # Forçar fechamento via JavaScript
                    driver.execute_script("""
                        if (typeof Swal !== 'undefined') {
                            Swal.close();
                        }
                        document.querySelectorAll('.swal2-container').forEach(el => el.remove());
                    """)
                    time.sleep(0.5)
                    log_message("✅ SweetAlert fechado via JavaScript", "INFO")

                return True
        except Exception:
            return False

    def processar_lote(self, driver, wait, exames_lote: list, modo_busca: str, cancel_flag, offset: int = 0):
        """Processa um lote de até 100 exames - registra falhas sem parar"""
        resultados_lote = []
        total_exames = len(exames_lote)

        for idx, exame in enumerate(exames_lote, start=1):
            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada pelo usuário.", "WARNING")
                break

            # Calcular posição global considerando lotes anteriores
            posicao_global = offset + idx
            log_message(f"➡️ Processando {modo_busca} [{posicao_global}/{offset + total_exames}]: {exame}", "INFO")

            try:
                # Aguardar página estar completamente carregada
                time.sleep(1)

                self.fechar_sweetalert(driver)

                # Verificar se há modais abertos e fechar
                try:
                    modal_backdrop = driver.find_element(By.CLASS_NAME, "modal-backdrop")
                    if modal_backdrop.is_displayed():
                        driver.execute_script("$('.modal').modal('hide');")
                        time.sleep(0.5)
                        log_message("🔄 Modal detectado e fechado", "INFO")
                except Exception:
                    pass

                campo_id = "numeroGuia" if modo_busca == "guia" else "numeroExame"
                log_message(f"🔍 Localizando campo de busca: {campo_id}", "INFO")

                # Estratégia robusta para localizar e interagir com o campo
                max_tentativas = 3
                campo_preenchido = False

                for tentativa in range(1, max_tentativas + 1):
                    try:
                        log_message(f"🔄 Tentativa {tentativa}/{max_tentativas} de preencher campo", "INFO")

                        # Aguardar elemento estar presente
                        campo_exame = wait.until(EC.presence_of_element_located((By.ID, campo_id)))

                        # Scroll até o elemento
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});",
                                              campo_exame)
                        time.sleep(0.5)

                        # Aguardar elemento estar visível
                        wait.until(EC.visibility_of_element_located((By.ID, campo_id)))

                        # Aguardar elemento estar clicável
                        wait.until(EC.element_to_be_clickable((By.ID, campo_id)))

                        # Remover atributos que podem bloquear interação
                        driver.execute_script("""
                            arguments[0].removeAttribute('readonly');
                            arguments[0].removeAttribute('disabled');
                            arguments[0].style.pointerEvents = 'auto';
                        """, campo_exame)

                        # Limpar campo usando múltiplas estratégias
                        try:
                            campo_exame.clear()
                            log_message("🧹 Campo limpo (método clear)", "INFO")
                        except Exception:
                            driver.execute_script("arguments[0].value = '';", campo_exame)
                            log_message("🧹 Campo limpo (JavaScript)", "INFO")

                        time.sleep(0.3)

                        # Tentar preencher campo
                        try:
                            campo_exame.send_keys(exame)
                            log_message(f"⌨️ Valor '{exame}' inserido (send_keys)", "INFO")
                            campo_preenchido = True
                        except Exception:
                            driver.execute_script("arguments[0].value = arguments[1];", campo_exame, exame)
                            # Disparar eventos para garantir que o valor seja reconhecido
                            driver.execute_script("""
                                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                            """, campo_exame)
                            log_message(f"⌨️ Valor '{exame}' inserido (JavaScript)", "INFO")
                            campo_preenchido = True

                        # Verificar se o valor foi realmente preenchido
                        valor_atual = driver.execute_script("return arguments[0].value;", campo_exame)
                        if valor_atual == exame:
                            log_message(f"✅ Campo preenchido corretamente: {valor_atual}", "SUCCESS")
                            break
                        else:
                            log_message(f"⚠️ Valor esperado '{exame}', obtido '{valor_atual}'", "WARNING")
                            if tentativa < max_tentativas:
                                time.sleep(1)
                                continue

                    except Exception as e:
                        log_message(f"⚠️ Erro na tentativa {tentativa}: {e}", "WARNING")
                        if tentativa < max_tentativas:
                            time.sleep(1)
                        else:
                            raise Exception(f"Falha ao preencher campo após {max_tentativas} tentativas")

                if not campo_preenchido:
                    raise Exception(f"Não foi possível preencher o campo {campo_id}")

                time.sleep(0.5)

                log_message("🔎 Clicando no botão de pesquisa...", "INFO")

                self.fechar_sweetalert(driver)

                try:
                    # Estratégia 1: Aguardar elemento estar clicável
                    botao_pesquisa = wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento")))

                    # Tentar clicar normalmente
                    try:
                        botao_pesquisa.click()
                        log_message("✅ Botão de pesquisa clicado (click normal)", "INFO")
                    except Exception as e:
                        log_message(f"⚠️ Click normal falhou: {e}. Tentando JavaScript...", "WARNING")

                        # Estratégia 2: Click via JavaScript
                        driver.execute_script("arguments[0].click();", botao_pesquisa)
                        log_message("✅ Botão de pesquisa clicado (JavaScript)", "INFO")

                except Exception as e:
                    log_message(f"⚠️ Erro ao clicar no botão. Tentando localizar novamente: {e}", "WARNING")

                    # Estratégia 3: Localizar novamente e usar JavaScript diretamente
                    time.sleep(1)
                    botao_retry = driver.find_element(By.ID, "pesquisaFaturamento")

                    # Remover atributo disabled se existir
                    driver.execute_script("arguments[0].removeAttribute('disabled');", botao_retry)

                    # Scroll e click
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_retry)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", botao_retry)
                    log_message("✅ Botão de pesquisa clicado (retry com JavaScript)", "INFO")

                try:
                    modal_carregando = driver.find_element(By.XPATH,
                                                           "//div[contains(@class,'modal-body') and contains(., 'Carregando')]")
                    if modal_carregando.is_displayed():
                        log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                        log_message("✅ Modal de carregamento fechado", "INFO")
                except Exception:
                    log_message("ℹ️ Modal não detectado. Prosseguindo...", "INFO")

                time.sleep(1)

                log_message("📋 Validando resultados da tabela...", "INFO")
                tbody_rows = driver.find_elements(By.CSS_SELECTOR, "#tabelaPreFaturamentoTbody tr")
                log_message(f"📊 Encontradas {len(tbody_rows)} linha(s) na tabela", "INFO")

                if len(tbody_rows) == 0:
                    log_message(f"⚠️ Nenhum resultado encontrado para {exame}. Pulando.", "WARNING")
                    resultados_lote.append({"exame": exame, "status": "sem_resultados"})
                    continue

                time.sleep(1)

                log_message("☑️ Marcando checkbox 'checkTodosPreFaturar'...", "INFO")

                self.fechar_sweetalert(driver)

                try:
                    log_message("☑️ Marcando checkbox (Método 1: JavaScript)...", "INFO")
                    # O script JS retorna true se o elemento for encontrado e clicado
                    clicked_with_js = driver.execute_script("""
                        const checkbox = document.getElementById('checkTodosPreFaturar');
                        if (checkbox) {
                            checkbox.click();
                            return true;
                        }
                        return false;
                    """)

                    if not clicked_with_js:
                        # Se o JS não encontrou o elemento, lança uma exceção para acionar o fallback
                        raise Exception("Checkbox 'checkTodosPreFaturar' não encontrado via JavaScript.")

                    log_message("✅ Checkbox marcado com sucesso (Método 1: JavaScript).", "INFO")

                except Exception as e:
                    log_message(f"⚠️ Método 1 (JS) falhou: {e}. Tentando fallback (Método 2: WebDriverWait)...",
                                "WARNING")
                    try:
                        # Fallback: Aguardar elemento estar clicável e clicar
                        checkbox = wait.until(EC.element_to_be_clickable((By.ID, "checkTodosPreFaturar")))
                        checkbox.click()
                        log_message("✅ Checkbox marcado com sucesso (Método 2: WebDriverWait).", "INFO")
                    except Exception as e2:
                        log_message(f"❌ Todas as tentativas de marcar o checkbox falharam: {e2}", "ERROR")
                        # Lança a exceção para que o processamento do exame seja interrompido e registrado como erro
                        raise Exception("Não foi possível marcar o checkbox 'checkTodosPreFaturar'.")

                # Aguardar modal de carregamento desaparecer após marcar checkbox
                log_message("⏳ Aguardando processamento após marcar checkbox...", "INFO")
                try:
                    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                    log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                    log_message("✅ Modal de carregamento fechado", "INFO")
                except Exception:
                    log_message("ℹ️ Modal não detectado. Prosseguindo...", "INFO")

                time.sleep(1)

                log_message("🎬 Clicando no botão 'Ações'...", "INFO")

                self.fechar_sweetalert(driver)
                time.sleep(1)

                # Laço para garantir que o status seja alterado para "Online"
                max_tentativas_status = 3
                status_alterado = False
                for tentativa_status in range(1, max_tentativas_status + 1):
                    log_message(f"🔄 Tentativa {tentativa_status}/{max_tentativas_status} para definir status como 'Online'", "INFO")

                    # Clicar no botão 'Ações'
                    try:
                        log_message("📡 Tentando clicar em 'Ações' (Método 1: JavaScript click)...", "INFO")
                        # O script JS localiza o botão via XPath e clica, retornando true se bem-sucedido.
                        clicked_with_js = driver.execute_script("""
                            const xpath = "//a[contains(@class, 'toggleMaisDeUm') and contains(., 'Ações')]";
                            const acoesBtn = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                            if (acoesBtn) {
                                acoesBtn.click();
                                return true;
                            }
                            return false;
                        """)

                        if not clicked_with_js:
                            raise Exception("Botão 'Ações' não encontrado ou clicado via JavaScript.")

                        log_message("✅ Botão 'Ações' clicado com sucesso (Método 1: JavaScript).", "INFO")
                        time.sleep(1)  # Aguardar o menu de ações abrir

                    except Exception as e1:
                        log_message(f"⚠️ Método 1 (JS) falhou: {e1}. Tentando fallback (Método 2: WebDriverWait)...",
                                    "WARNING")
                        try:
                            # Fallback para WebDriverWait + click
                            acoes_btn = wait.until(EC.element_to_be_clickable((
                                By.XPATH, "//a[contains(@class, 'toggleMaisDeUm') and contains(., 'Ações')]"
                            )))
                            acoes_btn.click()
                            log_message("✅ Botão 'Ações' clicado com sucesso (Método 2: WebDriverWait).", "INFO")
                            time.sleep(1)
                        except Exception as e2:
                            log_message(f"⚠️ Método 2 falhou: {e2}. Tentando fallback (Método 3: Scroll + JS)...",
                                        "WARNING")
                            try:
                                # Fallback final: Força o scroll e tenta o clique com JS
                                acoes_btn = driver.find_element(By.XPATH,
                                                                "//a[contains(@class, 'toggleMaisDeUm') and contains(., 'Ações')]")
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", acoes_btn)
                                time.sleep(0.5)
                                driver.execute_script("arguments[0].click();", acoes_btn)
                                log_message("✅ Botão 'Ações' clicado com sucesso (Método 3).", "INFO")
                                time.sleep(1)
                            except Exception as e3:
                                log_message(
                                    f"❌ Todas as tentativas de clicar em 'Ações' falharam na tentativa {tentativa_status}: {e3}",
                                    "ERROR")
                                time.sleep(1)
                                continue

                    # Clicar na opção 'Online' de forma mais robusta
                    try:
                        log_message("📡 Tentando clicar em 'On-line' com JavaScript (método primário)...", "INFO")
                        # O script JS retorna true se o botão for encontrado e clicado, senão false.
                        clicked_with_js = driver.execute_script("""
                            const onlineBtn = document.querySelector("a[data-url*='statusConferido=O']");
                            if (onlineBtn) {
                                onlineBtn.click();
                                return true;
                            }
                            return false;
                        """)

                        if not clicked_with_js:
                            # Se o JS não encontrou o botão, lança uma exceção para acionar o fallback.
                            raise Exception("Botão 'On-line' não encontrado via JavaScript.")

                        log_message("✅ Opção 'On-line' clicada com sucesso via JavaScript.", "INFO")

                    except Exception as e:
                        log_message(f"⚠️ Clique com JavaScript falhou: {e}. Usando fallback com WebDriverWait.",
                                    "WARNING")
                        # Fallback: Tenta o clique padrão com espera explícita.
                        online_btn = wait.until(EC.element_to_be_clickable((
                            By.CSS_SELECTOR, "a[data-url*='statusConferido=O']"
                        )))
                        online_btn.click()
                        log_message("✅ Opção 'On-line' clicada com sucesso via fallback (WebDriverWait).", "INFO")

                    # Aguardar o processamento (spinner desaparecer)
                    try:
                        spinner_wait = WebDriverWait(driver, 15)  # Aumentado para 15s
                        spinner_wait.until(EC.invisibility_of_element_located((By.ID, "spinner")))
                        log_message("✅ Processamento do status concluído (spinner desapareceu)", "INFO")
                    except Exception:
                        log_message("ℹ️ Spinner não detectado ou já invisível, aguardando tempo fixo.", "INFO")
                        time.sleep(1.5)

                    # Validação mais robusta, verificando o texto dentro da célula <td>
                    max_tentativas_validacao = 3
                    for tentativa_validacao in range(1, max_tentativas_validacao + 1):
                        try:
                            log_message(
                                f"🔎 Validando status... (Tentativa {tentativa_validacao}/{max_tentativas_validacao})",
                                "INFO")
                            seletor_celula_status = "#tabelaPreFaturamentoTbody tr:first-child td:nth-child(2)"

                            # Espera o texto "On-line" aparecer na célula (tempo de espera reduzido por tentativa)
                            WebDriverWait(driver, 5).until(EC.text_to_be_present_in_element(
                                (By.CSS_SELECTOR, seletor_celula_status), "On-line"
                            ))

                            log_message("✅ Validação bem-sucedida: Status é 'On-line'.", "SUCCESS")
                            status_alterado = True
                            break  # Sai do laço de validação

                        except TimeoutException:
                            log_message(
                                f"⚠️ Validação falhou na tentativa {tentativa_validacao}. O texto 'On-line' não apareceu a tempo.",
                                "WARNING")
                            if tentativa_validacao < max_tentativas_validacao:
                                time.sleep(2)  # Aguarda 2 segundos antes de tentar validar novamente
                            continue  # Próxima tentativa de validação

                        except Exception as e:
                            log_message(
                                f"⚠️ Erro inesperado ao validar o status na tabela (tentativa {tentativa_validacao}): {e}",
                                "WARNING")
                            continue  # Sai do laço de validação em caso de erro inesperado

                    if status_alterado:
                        break

                    # Se a validação falhou, garante que menus suspensos estejam fechados antes da próxima tentativa
                    driver.execute_script("document.body.click();")
                    time.sleep(1)

                if not status_alterado:
                    status_final = "Não foi possível ler"
                    try:
                        status_final = driver.find_element(By.CSS_SELECTOR, "#tabelaPreFaturamentoTbody tr:first-child td:nth-child(2)").text.strip()
                    except Exception:
                        pass

                    mensagem_erro = f"Não foi possível alterar o status para 'Online'. Status final encontrado: '{status_final}'."
                    log_message(f"❌ Falha ao alterar o status para 'Online' após {max_tentativas_status} tentativas.", "ERROR")
                    raise Exception(mensagem_erro)

                time.sleep(1)

                if modo_busca == "guia":
                    log_message("🔄 Modo guia detectado - Aguardando processamento adicional...", "INFO")
                    try:
                        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                        log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                        log_message("✅ Modal de carregamento fechado", "INFO")
                    except Exception:
                        log_message("ℹ️ Modal não detectado. Prosseguindo...", "INFO")
                        time.sleep(1)

                resultados_lote.append({"exame": exame, "status": "sucesso"})
                log_message(f"✅ {modo_busca.title()} {exame} processado com sucesso.", "SUCCESS")
            except Exception as e:
                log_message(f"❌ Erro ao processar o {modo_busca} '{exame}': {e}", "ERROR")
                resultados_lote.append({"exame": exame, "status": "erro", "detalhe": str(e)})
                continue

        return resultados_lote

    def gerar_ou_enviar_lote(self, driver, wait, gera_xml_tiss: str, username: str, password: str,
                             unimed_user: str, unimed_pass: str, pasta_download: str,
                             headless_mode: bool, cancel_flag, modo_busca: str, numero_lote: int, total_exames_lote: int):
        """Gera ou envia o lote após preparação"""
        if gera_xml_tiss == "sim":
            log_message(f"📤 Gerando e enviando XML para Unimed - Lote {numero_lote}...", "INFO")
            automacao = XMLGeneratorAutomation(username, password, pasta_download=pasta_download,headless=headless_mode)
            sucesso_envio = automacao.executar_processo_completo_login_navegacao(unimed_user, unimed_pass,cancel_flag=cancel_flag, total_exames_lote=total_exames_lote)
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

        url = os.getenv("SYSTEM_URL", "https://dap.pathoweb.com.br/login/auth")
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

            elif "moduloFaturamento" in current_url:
                log_message("✅ Já está no módulo de faturamento - pulando navegação", "SUCCESS")
            else:
                log_message(f"⚠️ URL inesperada detectada: {current_url}", "WARNING")
                driver.get("https://dap.pathoweb.com.br/moduloFaturamento/index")
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

            offset_global = 0
            # Processar cada lote sequencialmente
            for idx, lote in enumerate(lotes, start=1):
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break

                log_message(f"\n{'=' * 60}", "INFO")
                log_message(f"📦 PROCESSANDO LOTE {idx} de {total_lotes} ({len(lote)} exames)", "INFO")
                log_message(f"{'=' * 60}\n", "INFO")

                # Processar o lote atual
                resultados_lote = self.processar_lote(driver, wait, lote, modo_busca, cancel_flag, offset=offset_global)
                todos_resultados.extend(resultados_lote)

                offset_global += len(lote)

                # Resumo do lote atual
                sucesso_lote = [r for r in resultados_lote if r["status"] == "sucesso"]
                erro_lote = [r for r in resultados_lote if r["status"] == "erro"]
                sem_resultados_lote = [r for r in resultados_lote if r["status"] == "sem_resultados"]

                log_message(f"\n📊 Resumo do Lote {idx}:", "INFO")
                log_message(f"✅ Sucesso: {len(sucesso_lote)}", "SUCCESS")
                log_message(f"⚠️ Sem resultados: {len(sem_resultados_lote)}", "WARNING")
                log_message(f"❌ Erros: {len(erro_lote)}", "ERROR")
                if erro_lote:
                    exames_com_erro = [r['exame'] for r in erro_lote]
                    log_message(f"   - Exames com erro: {exames_com_erro}", "ERROR")

                # Gerar/enviar lote após processamento, passando a contagem de exames bem-sucedidos
                if len(sucesso_lote) > 0:
                    self.gerar_ou_enviar_lote(driver, wait, gera_xml_tiss, username, password,
                                              unimed_user, unimed_pass, pasta_download, headless_mode,
                                              cancel_flag, modo_busca, idx, total_exames_lote=len(sucesso_lote))
                else:
                    log_message(f"ℹ️ Nenhum exame processado com sucesso no lote {idx}. Pulando etapa de geração/envio.", "INFO")

                # Se não for o último lote, retornar à tela inicial
                # if idx < total_lotes:
                #     self.voltar_tela_inicial_preparacao(driver, wait)

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

            # Construir a mensagem final para o messagebox
            mensagem_final = (
                f"✅ Processamento finalizado!\n\n"
                f"📦 Lotes processados: {total_lotes}\n"
                f"📊 Total de exames: {total}\n"
                f"✅ Sucesso: {len(sucesso_total)}\n"
                f"⚠️ Sem resultados: {len(sem_resultados_total)}\n"
                f"❌ Erros: {len(erro_total)}"
            )

            # Adicionar a lista de exames com erro, se houver
            if erro_total:
                exames_com_erro_str = ", ".join([str(r['exame']) for r in erro_total])
                log_message(f"   - Exames com erro (final): {exames_com_erro_str}", "ERROR")
                mensagem_final += f"\n\nExames com erro:\n{exames_com_erro_str}"

            messagebox.showinfo("Sucesso", mensagem_final)

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            driver.quit()


def run(params: dict):
    module = PreparacaoLoteMultiploModule()
    module.run(params)