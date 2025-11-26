import os
import re
import time
import html
import pandas as pd
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from dotenv import load_dotenv
from datetime import datetime

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule
from src.modules.guias.lancamento_guia_unimed import LancamentoGuiaUnimedModule

load_dotenv()


class GuiaUnimedExamesModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Guia Unimed [Exames]")

    @staticmethod
    def get_unique_guias(file_path: str) -> list:
        try:
            df = pd.read_excel(file_path, header=0)
            guias = df.iloc[:, 0].dropna().tolist()
            if guias and isinstance(guias[0], str) and guias[0].upper() == "GUIA":
                guias = guias[1:]
            guias = [str(guia).strip() for guia in guias if str(guia).strip()]
            return guias
        except Exception as e:
            raise ValueError(f"Erro ao ler o Excel: {e}")

    def run(self, params: dict):
        username = params.get("username")
        password = params.get("password")
        unimed_user = params.get("unimed_user")
        unimed_pass = params.get("unimed_pass")
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode")
        excel_file = params.get("excel_file")

        login_url = os.getenv("SYSTEM_URL", "https://dap.pathoweb.com.br/login/auth")
        modulo_exame_url = "https://dap.pathoweb.com.br/moduloExame/index"

        driver = BrowserFactory.create_chrome(headless=headless_mode)
        wait = WebDriverWait(driver, 15)
        wait_long = WebDriverWait(driver, 30)

        try:
            log_message("Iniciando automação Guia Unimed [Exames]...", "INFO")

            if not unimed_user or not unimed_pass:
                messagebox.showerror("Erro", "Credenciais da Unimed são obrigatórias.")
                return

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

            guias_processadas_prev = self._carregar_guias_processadas_excel(excel_file)
            if guias_processadas_prev:
                log_message(
                    f"ℹ️ Encontradas {len(guias_processadas_prev)} guias já autorizadas no Excel (Status_Processamento=SUCESSO). "
                    "Pulando etapa de lançamento para elas.",
                    "INFO"
                )

            log_message(f"✅ Carregadas {len(guias)} guias do Excel", "SUCCESS")

            resultados_df = pd.DataFrame(columns=["GUIA", "CARTAO", "MEDICO", "CRM", "PROCEDIMENTOS", "QTD", "TEXTO"])
            resultados = []

            driver.get(login_url)
            wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
            driver.find_element(By.ID, "password").send_keys(password)
            driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

            self._navegar_para_modulo_exame(driver, wait, modulo_exame_url)

            for guia in guias:
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break

                if guia in guias_processadas_prev:
                    log_message(f"ℹ️ Guia {guia} já possui Numero_Guia e Status_Processamento=SUCESSO no Excel. "
                                "Pulando leitura no Pathoweb para esta guia.", "INFO")
                    resultados.append({"guia": guia, "status": "ja_processada"})
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

                try:
                    dados = self._processar_guia(driver, wait, wait_long, modulo_exame_url, guia)
                    resultados.append({"guia": guia, "status": "sucesso", "dados": dados})
                    resultados_df = pd.concat([resultados_df, pd.DataFrame([{
                        "GUIA": guia,
                        "CARTAO": dados.get("cartao", ""),
                        "MEDICO": dados.get("medico", ""),
                        "CRM": dados.get("crm", ""),
                        "PROCEDIMENTOS": dados.get("procedimentos", ""),
                        "QTD": dados.get("quantidades", ""),
                        "TEXTO": dados.get("texto", "")
                    }])], ignore_index=True)
                except Exception as e:
                    resultados.append({"guia": guia, "status": "erro", "erro": str(e)})
                    resultados_df = pd.concat([resultados_df, pd.DataFrame([{
                        "GUIA": guia,
                        "CARTAO": "",
                        "MEDICO": "",
                        "CRM": "",
                        "PROCEDIMENTOS": "",
                        "QTD": "",
                        "TEXTO": ""
                    }])], ignore_index=True)

            dados_para_lancamento = [
                r["dados"] for r in resultados
                if r.get("status") == "sucesso" and r.get("dados")
            ]

            output_file = self._salvar_resultados(resultados_df, excel_file)

            if output_file:
                atualizacoes_preexistentes = [
                    {"guia": guia, "numero_guia": numero, "status": "sucesso"}
                    for guia, numero in guias_processadas_prev.items()
                ]

                dados_para_lancamento = [
                    dados for dados in dados_para_lancamento
                    if dados.get("guia") not in guias_processadas_prev
                ]

                atualizacoes = list(atualizacoes_preexistentes)

                if dados_para_lancamento:
                    atualizacoes.extend(
                        self._lancar_guias_unimed(
                            dados_para_lancamento,
                            params,
                            cancel_flag,
                            headless_mode
                        )
                    )

                if atualizacoes:
                    self._atualizar_exames_pathoweb(
                        driver,
                        wait,
                        modulo_exame_url,
                        atualizacoes,
                        cancel_flag
                    )
            else:
                log_message("⚠️ Não foi possível salvar o arquivo de resultados; etapa de lançamento/atualização não será executada.", "WARNING")

            self._mostrar_resumo(resultados)

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            if not headless_mode:
                try:
                    input("Pressione Enter para fechar o navegador...")
                except EOFError:
                    pass
            driver.quit()

    def _navegar_para_modulo_exame(self, driver, wait, modulo_exame_url: str):
        log_message("Verificando módulo atual...", "INFO")
        time.sleep(2)
        current_url = driver.current_url

        if "moduloExame" in current_url:
            log_message("✅ Já está no módulo de exames.", "SUCCESS")
        else:
            trocar_modulo_url = "/site/trocarModulo?modulo=1"
            try:
                if "trocarModulo" in current_url or current_url.rstrip("/") == "https://dap.pathoweb.com.br":
                    log_message("Selecionando módulo de exames na tela de módulos...", "INFO")
                    link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"a[href='{trocar_modulo_url}']")))
                    link.click()
                    wait.until(EC.url_contains("moduloExame"))
                else:
                    raise TimeoutException("URL inesperada")
            except Exception:
                log_message("⚠️ Navegação direta para o módulo de exames (fallback)", "WARNING")
                driver.get(modulo_exame_url)
                wait.until(EC.url_contains("moduloExame"))

        self._fechar_modal(driver)
        driver.get(modulo_exame_url)
        log_message("✅ Módulo de exames carregado.", "SUCCESS")

    @staticmethod
    def _fechar_modal(driver):
        try:
            modal_close_button = driver.find_element(
                By.CSS_SELECTOR,
                "#mensagemParaClienteModal .modal-footer button"
            )
            if modal_close_button.is_displayed():
                modal_close_button.click()
                time.sleep(1)
        except Exception:
            pass

    def _processar_guia(self, driver, wait, wait_long, modulo_exame_url: str, guia: str) -> dict:
        log_message(f"➡️ Processando guia {guia}", "INFO")

        self._consultar_codigo_barras(driver, wait, modulo_exame_url, guia)
        self._clicar_botao_proximo(driver)
        self._aguardar_campos_detalhes(driver)

        cartao = self._obter_cartao(driver)
        medico = self._obter_medico(driver)
        crm = self._obter_crm(driver)
        procedimentos, quantidades = self._obter_procedimentos(driver)
        texto = self._obter_texto_clinico(driver)

        driver.execute_script("window.scrollTo(0, 0);")

        return {
            "guia": guia,
            "cartao": cartao or "",
            "medico": medico or "",
            "crm": crm or "",
            "procedimentos": ", ".join(procedimentos),
            "quantidades": ", ".join(quantidades),
            "texto": texto or ""
        }

    def _consultar_codigo_barras(self, driver, wait, modulo_exame_url: str, guia: str):
        input_field = None

        for tentativa in range(2):
            try:
                input_field = wait.until(EC.element_to_be_clickable((By.ID, "inputSearchCodBarra")))
                break
            except TimeoutException:
                if self._fechar_exame_se_aberto(driver):
                    continue
                log_message("⚠️ Campo de busca não encontrado, recarregando página...", "WARNING")
                driver.get(modulo_exame_url)

        if input_field is None:
            input_field = wait.until(EC.element_to_be_clickable((By.ID, "inputSearchCodBarra")))

        try:
            resultado_element = driver.find_element(By.CSS_SELECTOR, "#resultado")
            previous_html = resultado_element.get_attribute("innerHTML")
        except Exception:
            previous_html = ""

        input_field.clear()
        time.sleep(0.2)
        input_field.send_keys(str(guia))
        input_field.send_keys(Keys.ENTER)
        log_message("🔍 Enviando código para busca...", "INFO")

        def resultado_atualizado(drv):
            try:
                elem = drv.find_element(By.CSS_SELECTOR, "#resultado")
                return elem.get_attribute("innerHTML") != previous_html
            except Exception:
                return False

        try:
            WebDriverWait(driver, 20).until(resultado_atualizado)
        except TimeoutException:
            raise Exception("Consulta não retornou resultados a tempo.")

    def _clicar_botao_proximo(self, driver):
        log_message("➡️ Avançando para tela de detalhes do exame...", "INFO")

        self._esperar_formulario_paciente(driver)
        self._garantir_formulario_paciente_visivel(driver)

        selectors = [
            "a.btn.btn-sm.btn-primary.chamadaAjax.setupAjax[data-url='/paciente/saveAjax']",
            "a.btn.btn-primary.chamadaAjax.setupAjax[title='Próximo']",
            "a.chamadaAjax.btn.btn-primary"
        ]

        tentativas = 3
        for tentativa in range(1, tentativas + 1):
            botao = self._localizar_botao_proximo(driver, selectors)
            if not botao:
                log_message(f"⚠️ Botão 'Próximo' não localizado (tentativa {tentativa}/{tentativas})", "WARNING")
                time.sleep(1)
                self._garantir_formulario_paciente_visivel(driver)
                continue

            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)
                driver.execute_script("arguments[0].click();", botao)
                log_message(f"✅ Botão 'Próximo' clicado (tentativa {tentativa})", "SUCCESS")
            except Exception as e:
                log_message(f"⚠️ Erro ao clicar no botão 'Próximo': {e}", "WARNING")
                time.sleep(1)
                continue

            if self._aguardar_area_detalhes(driver):
                return

            log_message("⚠️ Detalhes do exame não aparecem, tentando novamente...", "WARNING")
            time.sleep(1)

        raise Exception("Não foi possível avançar para a tela de detalhes do exame.")

    @staticmethod
    def _localizar_botao_proximo(driver, selectors):
        for selector in selectors:
            try:
                return WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
            except Exception:
                continue
        return None

    @staticmethod
    def _garantir_formulario_paciente_visivel(driver):
        try:
            driver.execute_script("""
                const row = document.getElementById('rowPaciente');
                if (row) {
                    row.style.display = '';
                    if (typeof $ === 'function') {
                        $(row).show();
                    }
                }
                const container = document.getElementById('consultaCadastroPaciente');
                if (container) {
                    container.style.display = '';
                    if (typeof $ === 'function') {
                        $(container).show();
                    }
                }
            """)
        except Exception:
            pass

    def _esperar_formulario_paciente(self, driver):
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "cadastroPaciente")))
        except TimeoutException:
            log_message("⚠️ Formulário do paciente não apareceu dentro do tempo esperado.", "WARNING")

    def _aguardar_area_detalhes(self, driver):
        def detalhes_carregados(_):
            seletores = ["#requisicao_r", "#divSolicitacaoTA", "#divResultadoTA"]
            for sel in seletores:
                try:
                    elem = driver.find_element(By.CSS_SELECTOR, sel)
                    if elem.is_displayed():
                        return True
                except Exception:
                    continue
            return False

        try:
            WebDriverWait(driver, 20).until(detalhes_carregados)
            log_message("✅ Detalhes do exame carregados", "SUCCESS")
            return True
        except TimeoutException:
            log_message("⚠️ Detalhes do exame não apareceram a tempo", "WARNING")
            return False

    def _aguardar_campos_detalhes(self, driver):
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#requisicao_r #codigoUsuarioConvenio"))
            )
        except TimeoutException:
            log_message("⚠️ Campos do detalhe do exame demoraram para carregar.", "WARNING")

    def _obter_cartao(self, driver):
        try:
            cartao_input = driver.find_element(By.CSS_SELECTOR, "#requisicao_r #codigoUsuarioConvenio")
            cartao = cartao_input.get_attribute("value") or ""
            cartao = cartao.strip()
            if cartao:
                log_message(f"✅ Número do cartão obtido: {cartao}", "INFO")
                return cartao
        except Exception:
            pass

        try:
            cartao_anchor = driver.find_element(By.CSS_SELECTOR, "#codigoUsuarioConvenio + a.table-editable-ancora")
            cartao = cartao_anchor.text.strip()
            if cartao and cartao.lower() != "vazio":
                log_message(f"✅ Número do cartão obtido (âncora): {cartao}", "INFO")
                return cartao
        except Exception as e:
            log_message(f"⚠️ Erro ao obter cartão: {e}", "WARNING")
        return ""

    def _obter_medico(self, driver):
        medico = ""
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "medicoRequisitanteInput")))
            medico = driver.execute_script(
                "return $('#medicoRequisitanteInput').val && $('#medicoRequisitanteInput').val();")
            if medico and medico.strip():
                medico = medico.strip()
                log_message(f"✅ Médico requisitante encontrado (JS): {medico}", "SUCCESS")
                return medico
        except Exception:
            pass

        try:
            input_elem = driver.find_element(By.ID, "medicoRequisitanteInput")
            medico = input_elem.get_attribute("value").strip()
            if medico:
                log_message(f"✅ Médico requisitante encontrado (input): {medico}", "SUCCESS")
                return medico
        except Exception:
            pass

        try:
            ancora = driver.find_element(By.CSS_SELECTOR, "#requisicao_r a.table-editable-ancora.autocomplete.autocompleteSetup")
            medico = ancora.text.strip()
            if medico:
                log_message(f"✅ Médico requisitante encontrado (âncora): {medico}", "SUCCESS")
                return medico
        except Exception:
            pass

        try:
            label = driver.find_element(By.XPATH, "//td[contains(text(), 'Médico requisitante')]")
            medico_td = label.find_element(By.XPATH, "following-sibling::td")
            medico = medico_td.text.strip()
            if medico:
                log_message(f"✅ Médico requisitante encontrado (td): {medico}", "SUCCESS")
                return medico
        except Exception:
            pass

        log_message("⚠️ Não foi possível identificar o médico requisitante", "WARNING")
        return ""

    def _obter_crm(self, driver):
        log_message("Extraindo CRM do médico requisitante...", "INFO")

        try:
            input_elem = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#requisicao_r #medicoRequisitanteInput"))
            )
        except TimeoutException:
            log_message("⚠️ Campo de CRM não carregou a tempo.", "WARNING")
            return ""

        if not self._ativar_dropdown_medico(driver, input_elem):
            log_message("⚠️ Dropdown de CRM não abriu.", "WARNING")
            return ""

        try:
            dropdown_elem = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "ul.typeahead li.active a"))
            )
            crm_text = dropdown_elem.text
            crm_match = re.search(r'CRM:\s*(\S+)', crm_text)
            if crm_match:
                crm = crm_match.group(1)
                log_message(f"✅ CRM extraído: {crm}", "SUCCESS")
                driver.execute_script("document.body.click();")
                return crm
        except Exception as e:
            log_message(f"⚠️ Erro ao capturar CRM do dropdown: {e}", "WARNING")

        log_message("⚠️ CRM não encontrado", "WARNING")
        return ""

    def _lancar_guias_unimed(self, dados, params, cancel_flag, headless_mode):
        if not dados:
            return []

        dados_preparados = []
        for item in dados:
            dados_preparados.append({
                "guia": item.get("guia"),
                "cartao": item.get("cartao", ""),
                "medico": item.get("medico", ""),
                "crm": item.get("crm", ""),
                "procedimentos": item.get("procedimentos", ""),
                "qtd": item.get("quantidades", ""),
                "texto": item.get("texto", "")
            })

        resultado_atualizacoes = []
        lanc_mod = LancamentoGuiaUnimedModule()
        lanc_mod.headless_mode = headless_mode
        driver_unimed = BrowserFactory.create_chrome(headless=headless_mode)
        wait_unimed = WebDriverWait(driver_unimed, 15)

        try:
            lanc_mod.fazer_login_unimed(driver_unimed, wait_unimed, params.get("unimed_user"), params.get("unimed_pass"))
            lanc_mod.acessar_pagina_procedimento(driver_unimed)

            for idx, dados_guia in enumerate(dados_preparados, 1):
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada antes de finalizar os lançamentos na Unimed.", "WARNING")
                    break

                log_message(f"🚀 Lançando guia {dados_guia['guia']} na Unimed ({idx}/{len(dados)})", "INFO")
                try:
                    resultado = lanc_mod.processar_guia_unimed(driver_unimed, wait_unimed, dados_guia)
                    if resultado.get("status") in ["sucesso", "analise"] and resultado.get("numero_guia"):
                        resultado_atualizacoes.append({
                            "guia": dados_guia["guia"],
                            "numero_guia": resultado.get("numero_guia"),
                            "status": resultado.get("status")
                        })
                        log_message(f"✅ Guia {dados_guia['guia']} autorizada. Número: {resultado.get('numero_guia')}", "SUCCESS")
                    else:
                        log_message(f"❌ Falha ao lançar guia {dados_guia['guia']}: {resultado.get('erro')}", "ERROR")
                except Exception as e:
                    log_message(f"❌ Erro ao lançar guia {dados_guia['guia']}: {e}", "ERROR")

                if idx < len(dados):
                    lanc_mod.acessar_pagina_procedimento(driver_unimed)

        finally:
            driver_unimed.quit()

        return resultado_atualizacoes

    def _atualizar_exames_pathoweb(self, driver, wait, modulo_exame_url, atualizacoes, cancel_flag):
        if not atualizacoes:
            return

        log_message("📝 Atualizando dados das guias no PathoWeb (módulo de exames)...", "INFO")
        driver.get(modulo_exame_url)

        for idx, info in enumerate(atualizacoes, 1):
            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada antes de finalizar atualizações no PathoWeb.", "WARNING")
                break

            guia = info.get("guia")
            numero_guia = info.get("numero_guia")
            if not numero_guia:
                log_message(f"⚠️ Número da guia não disponível para {guia}, pulando atualização.", "WARNING")
                continue

            try:
                log_message(f"🔍 Atualizando exame {guia} ({idx}/{len(atualizacoes)})", "INFO")
                self._consultar_codigo_barras(driver, wait, modulo_exame_url, guia)
                self._clicar_botao_proximo(driver)
                self._aguardar_campos_detalhes(driver)
                self._preencher_dados_guia_no_exame(driver, numero_guia)
                self._clicar_salvar_dados_exame(driver)
                log_message(f"✅ Exame {guia} atualizado com número {numero_guia}", "SUCCESS")
            except Exception as e:
                log_message(f"❌ Erro ao atualizar exame {guia}: {e}", "ERROR")

    def _preencher_dados_guia_no_exame(self, driver, numero_guia: str):
        data_atual = datetime.now()
        data_ymd = data_atual.strftime("%Y-%m-%d")
        data_br = data_atual.strftime("%d/%m/%Y")

        script = f"""
        (function(){{
            function setValor(selector, valor, texto){{
                var input = document.querySelector(selector);
                if (!input) return;
                input.value = valor;
                input.setAttribute('value', valor);
                ['input','change','blur'].forEach(function(evt){{
                    var event = new Event(evt, {{ bubbles: true }});
                    input.dispatchEvent(event);
                }});
                var anchor = input.closest('td') ? input.closest('td').querySelector('a.table-editable-ancora') : null;
                if (anchor){{
                    anchor.textContent = texto || valor;
                    anchor.style.display = 'inline';
                }}
            }}

            setValor("#requisicao_r input[name='dataAutorizacao']", "{data_ymd}", "{data_br}");
            setValor("#requisicao_r input[name='dataRequisicao']", "{data_ymd}", "{data_br}");

            var numeroInput = document.getElementById('numeroGuiaInput');
            if (numeroInput){{
                numeroInput.value = "{numero_guia}";
                numeroInput.setAttribute('value', "{numero_guia}");
                ['input','change','blur','keyup'].forEach(function(evt){{
                    var event = new Event(evt, {{ bubbles: true }});
                    numeroInput.dispatchEvent(event);
                }});
                var anchor = numeroInput.closest('td') ? numeroInput.closest('td').querySelector('a.table-editable-ancora') : null;
                if (anchor){{
                    anchor.textContent = "{numero_guia}";
                    anchor.style.display = 'inline';
                }}
            }}
        }})();
        """
        driver.execute_script(script)

    def _clicar_salvar_dados_exame(self, driver):
        try:
            salvar_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "a.btn.btn-sm.btn-success.chamadaAjax.noValidate.setupAjax[data-url='/moduloExame/saveExameAjax'], "
                    "a#btnSaveAjaxNovalidate"
                ))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", salvar_btn)
            driver.execute_script("arguments[0].click();", salvar_btn)
            try:
                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                WebDriverWait(driver, 20).until(EC.invisibility_of_element_located((By.ID, "spinner")))
            except Exception:
                time.sleep(2)
        except Exception as e:
            raise Exception(f"Erro ao clicar em Salvar no módulo de exames: {e}")

    def _fechar_exame_se_aberto(self, driver) -> bool:
        try:
            fechar_btn = driver.find_element(By.ID, "fecharExameBarraFerramenta")
            if fechar_btn.is_displayed() and fechar_btn.is_enabled():
                log_message("ℹ️ Exame já estava aberto. Fechando para voltar ao campo de busca...", "INFO")
                driver.execute_script("arguments[0].click();", fechar_btn)
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "inputSearchCodBarra"))
                )
                log_message("✅ Exame fechado, campo de busca disponível novamente.", "SUCCESS")
                return True
        except Exception:
            pass
        return False

    def _carregar_guias_processadas_excel(self, excel_file: str) -> dict:
        try:
            df = pd.read_excel(excel_file, header=0)
            if df.empty:
                return {}
            df.columns = df.columns.str.upper().str.strip()
            if "GUIA" not in df.columns or "NUMERO_GUIA" not in df.columns or "STATUS_PROCESSAMENTO" not in df.columns:
                return {}

            gui_map = {}
            for _, row in df.iterrows():
                guia = str(row["GUIA"]).strip() if pd.notna(row["GUIA"]) else ""
                numero = str(row["NUMERO_GUIA"]).strip() if pd.notna(row["NUMERO_GUIA"]) else ""
                status = str(row["STATUS_PROCESSAMENTO"]).strip().upper() if pd.notna(row["STATUS_PROCESSAMENTO"]) else ""
                if guia and numero and status == "SUCESSO":
                    gui_map[guia] = numero
            return gui_map
        except Exception as e:
            log_message(f"⚠️ Não foi possível verificar guias já processadas: {e}", "WARNING")
            return {}

    @staticmethod
    def _ativar_dropdown_medico(driver, input_elem):
        def dropdown_visivel(_):
            try:
                dropdown = driver.find_element(By.CSS_SELECTOR, "ul.typeahead")
                return dropdown.is_displayed()
            except Exception:
                return False

        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_elem)
            input_elem.click()
            if WebDriverWait(driver, 3).until(dropdown_visivel):
                return True
        except Exception:
            pass

        try:
            ancora = driver.find_element(By.CSS_SELECTOR, "#requisicao_r a.table-editable-ancora.autocomplete.autocompleteSetup")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", ancora)
            ancora.click()
            if WebDriverWait(driver, 3).until(dropdown_visivel):
                return True
        except Exception:
            pass

        return False

    def _obter_procedimentos(self, driver):
        procedimentos = []
        quantidades = []

        try:
            linhas_procedimentos = driver.find_elements(
                By.CSS_SELECTOR,
                "#procedimentosForm tbody tr[id^='procedimento_']"
            )
            for linha in linhas_procedimentos:
                try:
                    checkbox = linha.find_element(By.CSS_SELECTOR, "input[name='procedimentoExameId']")
                    if checkbox.get_attribute("value") == "":
                        continue  # ignora linha vazia/novo procedimento
                except Exception:
                    continue

                try:
                    qtd_elem = linha.find_element(By.CSS_SELECTOR, "td:nth-child(2) .table-editable-ancora")
                    quantidade = qtd_elem.text.strip()
                except Exception:
                    quantidade = "1"

                try:
                    proc_elem = linha.find_element(By.CSS_SELECTOR, "td:nth-child(3) .table-editable-ancora")
                    proc_text = proc_elem.text.strip()
                    codigo = proc_text.split(" - ")[0].strip() if " - " in proc_text else proc_text
                except Exception:
                    codigo = ""

                if codigo:
                    procedimentos.append(codigo)
                    quantidades.append(quantidade if quantidade else "1")
        except Exception as e:
            log_message(f"⚠️ Erro ao coletar procedimentos: {e}", "WARNING")

        if not procedimentos:
            log_message("⚠️ Nenhum procedimento encontrado na tabela principal, tentando fallback...", "WARNING")
            procedimentos, quantidades = self._extrair_dados_genericos(driver)

        if procedimentos:
            log_message(f"✅ Códigos de procedimentos obtidos: {', '.join(procedimentos)}", "INFO")
            log_message(f"✅ Quantidades obtidas: {', '.join(quantidades)}", "INFO")
        else:
            log_message("⚠️ Nenhum procedimento encontrado", "WARNING")

        return procedimentos, quantidades if quantidades else []

    def _extrair_dados_genericos(self, driver):
        procedimentos = []
        quantidades = []
        selectors = [
            "#divSolicitacaoTA table tbody tr",
            "#divResultadoTA table tbody tr",
            "#resultado table tbody tr",
            ".table-responsive table tbody tr",
            "table.table tbody tr"
        ]

        for selector in selectors:
            try:
                rows = driver.find_elements(By.CSS_SELECTOR, selector)
                rows = [row for row in rows if row.is_displayed()]
                if not rows:
                    continue
                for row in rows:
                    cells = row.find_elements(By.CSS_SELECTOR, "td")
                    if not cells:
                        continue
                    procedimento = cells[-2].text.strip() if len(cells) >= 2 else ""
                    quantidade = cells[-1].text.strip() if cells else "1"
                    if procedimento:
                        procedimentos.append(procedimento)
                        quantidades.append(quantidade if quantidade else "1")
                if procedimentos:
                    break
            except Exception:
                continue

        return procedimentos, quantidades

    def _obter_texto_clinico(self, driver):
        texto = ""
        try:
            hidden = driver.find_element(By.ID, "conteudo_dadosClinicosTexto")
            hidden_value = html.unescape(hidden.get_attribute("value") or "").strip()
            if hidden_value:
                log_message(f"✅ Texto clínico obtido de campo oculto: {hidden_value[:60]}...", "INFO")
                return hidden_value
        except Exception:
            pass

        try:
            textarea = driver.find_element(By.ID, "dadosClinicosTexto")
            textarea_value = html.unescape(textarea.get_attribute("value") or "").strip()
            if textarea_value:
                log_message(f"✅ Texto clínico obtido do textarea: {textarea_value[:60]}...", "INFO")
                return textarea_value
        except Exception:
            pass
        try:
            iframe = driver.find_element(By.CSS_SELECTOR, "#myModal .cke_wysiwyg_frame")
            driver.switch_to.frame(iframe)
            texto_element = driver.find_element(By.CSS_SELECTOR, "body")
            texto = texto_element.text.strip()
            driver.switch_to.default_content()
            if texto:
                log_message(f"✅ Texto clínico obtido via iframe: {texto[:60]}...", "INFO")
                return texto
        except Exception:
            driver.switch_to.default_content()

        try:
            texto_element = driver.find_element(
                By.XPATH,
                "//div[contains(@class, 'form-group')]//*[contains(text(), 'Dados clínicos')]/following-sibling::*"
            )
            texto = texto_element.text.strip()
            if texto:
                log_message(f"✅ Texto clínico obtido via XPath: {texto[:60]}...", "INFO")
                return texto
        except Exception:
            pass

        try:
            elements = driver.find_elements(By.CSS_SELECTOR, "#myModal div.form-control, #myModal textarea.form-control")
            for elem in elements:
                if elem.text and len(elem.text.strip()) > 5:
                    texto = elem.text.strip()
                    log_message(f"✅ Texto clínico obtido via fallback: {texto[:60]}...", "INFO")
                    return texto
        except Exception:
            pass

        log_message("⚠️ Texto clínico não localizado", "WARNING")
        return ""

    def _salvar_resultados(self, resultados_df: pd.DataFrame, excel_file: str):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = os.path.dirname(excel_file)
            output_file = os.path.join(output_dir, f"resultados_guias_unimed_exames_{timestamp}.xlsx")
            resultados_df.to_excel(output_file, index=False)
            log_message(f"✅ Resultados salvos em: {output_file}", "SUCCESS")
            return output_file
        except Exception as e:
            log_message(f"❌ Erro ao salvar resultados: {e}", "ERROR")
            return None

    def _mostrar_resumo(self, resultados: list):
        total = len(resultados)
        sucesso = [r for r in resultados if r["status"] == "sucesso"]
        erro = [r for r in resultados if r["status"] == "erro"]

        log_message("\nResumo do processamento:", "INFO")
        log_message(f"Total de guias: {total}", "INFO")
        log_message(f"Processadas com sucesso: {len(sucesso)}", "SUCCESS")
        log_message(f"Erros: {len(erro)}", "ERROR")

        messagebox.showinfo(
            "Resultado",
            f"✅ Processamento finalizado!\n"
            f"Total: {total}\n"
            f"Sucesso: {len(sucesso)}\n"
            f"Erros: {len(erro)}"
        )


def run(params: dict):
    module = GuiaUnimedExamesModule()
    module.run(params)

