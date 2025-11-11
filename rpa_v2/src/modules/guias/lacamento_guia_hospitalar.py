import os
import time
import re
import difflib
import pandas as pd
from datetime import datetime
from tkinter import messagebox

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule


class LancamentoGuiaHospitalarModule(BaseModule):
    LOGIN_URL = "https://servico.planohospitalar.org.br/solusweb/prestador/index.php"
    MENU_NAV_SELECTOR = "div.nav-collapse"
    MODAL_SELECTOR = "div.modal.in, div.modal.show"

    def __init__(self):
        super().__init__(nome="Lançamento Guia Hospitalar")
        self.headless_mode = False

    def click_element(self, driver, element, descricao="elemento"):
        try:
            if self.headless_mode:
                driver.execute_script("arguments[0].click();", element)
                log_message(f"✅ Clique via JavaScript em {descricao}", "INFO")
            else:
                try:
                    element.click()
                    log_message(f"✅ Clique normal em {descricao}", "INFO")
                except Exception:
                    driver.execute_script("arguments[0].click();", element)
                    log_message(f"✅ Clique via JavaScript (fallback) em {descricao}", "INFO")
        except Exception as e:
            log_message(f"❌ Erro ao clicar em {descricao}: {e}", "ERROR")
            raise

    def wait_for_element(self, driver, wait, by, value, condition="presence", timeout=None):
        try:
            if timeout:
                wait = WebDriverWait(driver, timeout)

            if self.headless_mode and condition in ["clickable", "visible"]:
                element = wait.until(EC.presence_of_element_located((by, value)))
            elif condition == "clickable":
                element = wait.until(EC.element_to_be_clickable((by, value)))
            elif condition == "visible":
                element = wait.until(EC.visibility_of_element_located((by, value)))
            else:
                element = wait.until(EC.presence_of_element_located((by, value)))
            return element
        except Exception as e:
            log_message(f"❌ Erro ao aguardar elemento {value}: {e}", "ERROR")
            raise

    def set_input_value(self, driver, element, valor, descricao):
        try:
            if element.is_displayed():
                driver.execute_script("arguments[0].focus();", element)
                element.clear()
                if valor:
                    element.send_keys(valor)
                log_message(f"📝 Digitado '{valor}' em {descricao}", "INFO")
            else:
                raise Exception("elemento não visível")
        except Exception as e:
            log_message(f"⚠️ Falha ao digitar em {descricao}: {e} - usando JavaScript", "WARNING")
            driver.execute_script(
                """
                const campo = arguments[0];
                const valor = arguments[1];
                campo.value = valor;
                campo.setAttribute('value', valor);
                campo.dispatchEvent(new Event('input', { bubbles: true }));
                campo.dispatchEvent(new Event('change', { bubbles: true }));
                campo.dispatchEvent(new Event('blur', { bubbles: true }));
                """,
                element,
                valor
            )
            time.sleep(0.2)

    def read_excel_data(self, file_path: str) -> list:
        try:
            df = pd.read_excel(file_path, header=0)
            expected_columns = ['GUIA', 'CARTAO', 'MEDICO', 'CRM', 'PROCEDIMENTOS', 'QTD', 'TEXTO']
            df.columns = df.columns.str.upper().str.strip()
            missing_columns = [col for col in expected_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Colunas faltando no Excel: {missing_columns}")

            data_list = []
            for _, row in df.iterrows():
                if pd.notna(row['GUIA']) and str(row['GUIA']).strip():
                    def converter_valor(valor):
                        if pd.notna(valor):
                            valor_str = str(valor).strip()
                            if valor_str.endswith('.0'):
                                valor_str = valor_str[:-2]
                            return valor_str
                        return ''

                    data_list.append({
                        'guia': converter_valor(row['GUIA']),
                        'cartao': converter_valor(row['CARTAO']),
                        'medico': converter_valor(row['MEDICO']),
                        'crm': converter_valor(row['CRM']),
                        'procedimentos': converter_valor(row['PROCEDIMENTOS']),
                        'qtd': converter_valor(row['QTD']),
                        'texto': converter_valor(row['TEXTO'])
                    })
            return data_list
        except Exception as e:
            raise ValueError(f"Erro ao ler o Excel: {e}")

    def comparar_nomes_medicos(self, nome_procurado, nome_encontrado):
        def normalizar(nome):
            return re.sub(r'\W+', '', nome.lower().strip())

        matcher = difflib.SequenceMatcher(
            None,
            normalizar(nome_procurado),
            normalizar(nome_encontrado)
        )
        return matcher.ratio()

    def extrair_apenas_numeros(self, valor):
        return re.sub(r'[^0-9]', '', str(valor)) if valor else ''

    def safe_switch_to_window(self, driver, handle):
        try:
            driver.switch_to.window(handle)
            return True
        except Exception as e:
            log_message(f"⚠️ Falha ao alternar para janela {handle}: {e}", "WARNING")
            return False

    def fazer_login(self, driver, wait, username, password):
        log_message("🔐 Fazendo login no portal Hospitalar...", "INFO")
        driver.get(self.LOGIN_URL)

        # Garantir que o formulário de login está visível antes de interagir
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "form#form1.form-signin")))

        campo_usuario = wait.until(EC.element_to_be_clickable((By.ID, "operador")))
        campo_usuario.clear()
        campo_usuario.send_keys(username)

        campo_senha = wait.until(EC.element_to_be_clickable((By.ID, "senha")))
        campo_senha.clear()
        campo_senha.send_keys(password)

        botao_entrar = wait.until(EC.element_to_be_clickable((By.ID, "entrar")))
        self.click_element(driver, botao_entrar, "botão Entrar")

        # Aguardar carregamento completo do menu principal após o login
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.nav-collapse")))
        wait.until(EC.visibility_of_element_located((By.XPATH, "//a[contains(., 'Emissão de guias')]")))
        time.sleep(1.5)
        log_message("✅ Login realizado com sucesso", "SUCCESS")

    def fechar_modal_pendente(self, driver):
        try:
            modais = driver.find_elements(By.CSS_SELECTOR, self.MODAL_SELECTOR)
            for modal in modais:
                if modal.is_displayed():
                    log_message("ℹ️ Modal pós-login detectado, tentando fechar...", "INFO")
                    botoes = modal.find_elements(By.CSS_SELECTOR, "button.btn")
                    for botao in botoes:
                        texto = botao.text.strip().lower()
                        if texto in ["ok", "fechar"]:
                            self.click_element(driver, botao, f"botão {texto.upper()} do modal")
                            time.sleep(1)
                            return
                    driver.execute_script("arguments[0].style.display='none';", modal)
                    time.sleep(1)
        except Exception as e:
            log_message(f"⚠️ Erro ao fechar modal: {e}", "WARNING")

    def navegar_para_guia_procedimento(self, driver, wait):
        self.fechar_modal_pendente(driver)

        log_message("🔍 Abrindo menu Emissão de guias...", "INFO")
        menu_emissao = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@class, 'dropdown-toggle') and contains(text(), 'Emissão de guias')]")
        ))
        self.click_element(driver, menu_emissao, "menu Emissão de guias")
        time.sleep(1)

        log_message("🔍 Acessando item Guia de SP/SADT...", "INFO")
        item_guia = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//a[contains(@href, 'prestador/procedimento.php') and contains(text(), 'Guia de SP/SADT')]"
        )))
        href_destino = item_guia.get_attribute("href")
        log_message(f"ℹ️ URL detectada para Guia SP/SADT: {href_destino}", "INFO")

        self.click_element(driver, item_guia, "Guia de SP/SADT")
        time.sleep(3)
        log_message("✅ Página de procedimentos acessada via menu", "SUCCESS")

    def preencher_codigo_beneficiario(self, driver, wait, cartao):
        cartao_formatado = str(cartao).strip()
        campo_codigo = self.wait_for_element(driver, wait, By.ID, "codigo", condition="presence")
        campo_codigo.clear()
        campo_codigo.send_keys(cartao_formatado)
        log_message(f"✅ Código do beneficiário preenchido: {cartao_formatado}", "SUCCESS")
        time.sleep(1)

        tentativas = driver.find_elements(By.CSS_SELECTOR, ".alert.alert-danger, #msg_erro")
        for aviso in tentativas:
            if aviso.is_displayed():
                texto = aviso.text.strip()
                if texto:
                    log_message(f"⚠️ Aviso após preencher beneficiário: {texto}", "WARNING")
                    return {'erro': True, 'mensagem': texto}
        return {'erro': False}

    def buscar_medico_solicitante(self, driver, wait, nome_medico, crm=None):
        janela_original = driver.current_window_handle
        log_message("🔍 Abrindo busca de solicitante...", "INFO")
        botao_busca = self.wait_for_element(driver, wait, By.ID, "busca_solicitante", condition="clickable")
        handles_antes = set(driver.window_handles)
        self.click_element(driver, botao_busca, "botão buscar solicitante")
        time.sleep(2)

        try:
            WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(handles_antes) or len(d.window_handles) > 1)
        except Exception:
            pass

        todas_janelas = driver.window_handles
        nova_janela = None
        for handle in todas_janelas:
            if handle not in handles_antes:
                nova_janela = handle
                break

        if nova_janela:
            if self.safe_switch_to_window(driver, nova_janela):
                log_message("✅ Popup do solicitante aberto em nova janela", "SUCCESS")
        elif len(todas_janelas) > 1:
            for handle in todas_janelas:
                if handle != janela_original:
                    if self.safe_switch_to_window(driver, handle):
                        log_message("✅ Popup do solicitante aberto (handle alternativo)", "SUCCESS")
                        break
                    break
        else:
            log_message("✅ Busca de solicitante aberta na mesma janela", "INFO")

        popup_frames = driver.find_elements(By.CSS_SELECTOR, "iframe#iframeCorpoClinico, iframe[name='iframetxt'], iframe[name='iframeCorpoClinico']")
        if popup_frames:
            driver.switch_to.frame(popup_frames[0])
            log_message("ℹ️ Popup com iframe detectado - alternando para o frame", "INFO")
        else:
            log_message("ℹ️ Popup sem iframe - usando DOM principal", "INFO")

        self.wait_for_element(driver, wait, By.CSS_SELECTOR, "form#form1", condition="presence")

        crm_numeros = self.extrair_apenas_numeros(crm)
        tentativas = [
            {'crm': crm_numeros, 'nome': nome_medico, 'log': 'CRM + Nome'},
            {'crm': crm_numeros, 'nome': '', 'log': 'Apenas CRM'},
            {'crm': '', 'nome': nome_medico, 'log': 'Apenas Nome'}
        ]

        medico_encontrado = None
        nome_resultado = ""

        for tentativa in tentativas:
            campo_nome = self.wait_for_element(driver, wait, By.ID, "nome", condition="presence")
            campo_conselho = self.wait_for_element(driver, wait, By.ID, "conselho", condition="presence")
            botao_localizar = self.wait_for_element(driver, wait, By.ID, "localizar", condition="clickable")

            self.set_input_value(driver, campo_nome, "", "campo Nome (limpar)")
            self.set_input_value(driver, campo_conselho, "", "campo Conselho (limpar)")

            if tentativa['crm']:
                self.set_input_value(driver, campo_conselho, tentativa['crm'], "campo Conselho (CRM)")

            if tentativa['nome']:
                self.set_input_value(driver, campo_nome, tentativa['nome'].upper(), "campo Nome")

            self.click_element(driver, botao_localizar, f"botão Localizar ({tentativa['log']})")
            time.sleep(3)

            try:
                tabela = self.wait_for_element(driver, wait, By.CSS_SELECTOR, "table.table-hover tbody", condition="presence", timeout=5)
                linhas = tabela.find_elements(By.TAG_NAME, "tr")
            except Exception:
                linhas = []

            candidatos = []
            for linha in linhas:
                colunas = linha.find_elements(By.TAG_NAME, "td")
                if len(colunas) >= 2:
                    nome_encontrado = colunas[1].text.strip()
                    if not nome_encontrado:
                        continue
                    similaridade = self.comparar_nomes_medicos(nome_medico, nome_encontrado)
                    documento = colunas[2].text.strip() if len(colunas) >= 3 else ""
                    log_message(f"📋 Resultado: {nome_encontrado} ({documento}) - Similaridade: {similaridade:.2f}", "INFO")
                    candidatos.append((similaridade, linha, nome_encontrado))

            if candidatos:
                candidatos.sort(key=lambda item: item[0], reverse=True)
                melhor_similaridade, linha_melhor, nome_resultado = candidatos[0]
                if melhor_similaridade >= 0.5 or not tentativa['nome']:
                    medico_encontrado = linha_melhor
                    break

        if not medico_encontrado:
            raise Exception(f"Médico não encontrado: {nome_medico}")

        self.click_element(driver, medico_encontrado, f"linha médico {nome_resultado}")
        log_message(f"✅ Médico selecionado: {nome_resultado}", "SUCCESS")

        retornou = False
        for _ in range(20):
            handles_atuais = driver.window_handles
            if janela_original in handles_atuais:
                if self.safe_switch_to_window(driver, janela_original):
                    retornou = True
                    break
            time.sleep(0.3)

        if not retornou:
            handles_atuais = driver.window_handles
            if handles_atuais:
                retornou = self.safe_switch_to_window(driver, handles_atuais[0])

        if not retornou:
            raise Exception("Não foi possível retornar à janela principal após selecionar médico")

        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        log_message("↩️ Retornou para janela principal", "INFO")

        try:
            WebDriverWait(driver, 10).until(
                EC.text_to_be_present_in_element_value((By.ID, "nome_solicitante"), nome_resultado.split('-')[0].strip())
            )
        except Exception:
            log_message("ℹ️ Campo nome_solicitante ainda não atualizado após 10s - prosseguindo mesmo assim", "INFO")
        time.sleep(1)

    def preencher_procedimentos(self, driver, procedimentos_str, quantidades_str):
        if not procedimentos_str or not quantidades_str:
            log_message("⚠️ Procedimentos ou quantidades vazios", "WARNING")
            return

        procedimentos = [p.strip() for p in str(procedimentos_str).split(',')]
        quantidades = [q.strip() for q in str(quantidades_str).split(',')]

        if len(procedimentos) != len(quantidades):
            log_message("⚠️ Quantidade de procedimentos difere da quantidade informada", "WARNING")
            tamanho = min(len(procedimentos), len(quantidades))
            procedimentos = procedimentos[:tamanho]
            quantidades = quantidades[:tamanho]

        for idx, (proc, qtd) in enumerate(zip(procedimentos, quantidades)):
            if idx >= 5:
                log_message("⚠️ Limite de 5 procedimentos atingido", "WARNING")
                break

            log_message(f"📝 Preenchendo procedimento {idx}: {proc} (qtd {qtd})", "INFO")
            try:
                campo_proc = self.wait_for_element(driver, WebDriverWait(driver, 10), By.ID, f"procedimento{idx}", condition="clickable")
                campo_proc.clear()
                campo_proc.send_keys(proc)
                campo_proc.send_keys(Keys.TAB)
                time.sleep(0.5)
            except Exception as e:
                log_message(f"⚠️ Falha ao digitar procedimento {idx}: {e} - usando JavaScript", "WARNING")
                driver.execute_script(
                    """
                    const campo = document.getElementById(arguments[0]);
                    if (campo) {
                        campo.value = arguments[1];
                        campo.dispatchEvent(new Event('change', { bubbles: true }));
                        campo.dispatchEvent(new Event('blur', { bubbles: true }));
                    }
                    """,
                    f"procedimento{idx}",
                    proc
                )

            try:
                campo_qtd = self.wait_for_element(driver, WebDriverWait(driver, 10), By.ID, f"quantidade{idx}", condition="clickable")
                campo_qtd.clear()
                campo_qtd.send_keys(qtd)
                campo_qtd.send_keys(Keys.TAB)
                time.sleep(0.5)
            except Exception as e:
                log_message(f"⚠️ Falha ao digitar quantidade {idx}: {e} - usando JavaScript", "WARNING")
                driver.execute_script(
                    """
                    const campo = document.getElementById(arguments[0]);
                    if (campo) {
                        campo.value = arguments[1];
                        campo.dispatchEvent(new Event('change', { bubbles: true }));
                        campo.dispatchEvent(new Event('blur', { bubbles: true }));
                    }
                    """,
                    f"quantidade{idx}",
                    qtd
                )

            log_message(f"✅ Procedimento {proc} preenchido", "SUCCESS")

    def preencher_campos_fixos(self, driver, wait):
        try:
            log_message("🛠️ Preenchendo campos fixos (regime, tipo, executante)...", "INFO")

            campo_regime = self.wait_for_element(driver, wait, By.ID, "regime_atendimento", condition="presence")
            try:
                Select(campo_regime).select_by_value("01")
            except Exception:
                driver.execute_script(
                    """
                    const select = arguments[0];
                    select.value = "01";
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    select.dispatchEvent(new Event('blur', { bubbles: true }));
                    """,
                    campo_regime
                )
            log_message("✅ Regime de atendimento definido para '01 - Ambulatorial'", "SUCCESS")

            campo_tipo = self.wait_for_element(driver, wait, By.ID, "tipo_atendimento", condition="presence")
            try:
                Select(campo_tipo).select_by_value("23")
            except Exception:
                driver.execute_script(
                    """
                    const select = arguments[0];
                    select.value = "23";
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    select.dispatchEvent(new Event('blur', { bubbles: true }));
                    """,
                    campo_tipo
                )
            log_message("✅ Tipo de atendimento definido para '23 - Exames'", "SUCCESS")

            campo_executante = self.wait_for_element(driver, wait, By.ID, "executante", condition="presence")
            try:
                Select(campo_executante).select_by_value("33119687")
            except Exception:
                driver.execute_script(
                    """
                    const select = arguments[0];
                    select.value = "33119687";
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    select.dispatchEvent(new Event('blur', { bubbles: true }));
                    """,
                    campo_executante
                )
            log_message("✅ Executante definido para 'DAP DIAGNOSTICO EM ANATOMIA PATOLOGICA...'", "SUCCESS")

            time.sleep(1)
        except Exception as e:
            log_message(f"⚠️ Erro ao preencher campos fixos: {e}", "WARNING")

    def fazer_login_pathoweb(self, driver, wait, username, password):
        try:
            log_message("🔐 Fazendo login no PathoWeb...", "INFO")
            driver.get("https://pathoweb.com.br/login/auth")

            wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
            driver.find_element(By.ID, "password").send_keys(password)
            botao_submit = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            self.click_element(driver, botao_submit, "botão login PathoWeb")

            log_message("Verificando módulo de faturamento...", "INFO")
            current_url = driver.current_url
            if current_url == "https://pathoweb.com.br/" or "trocarModulo" in current_url:
                try:
                    modulo_link = self.wait_for_element(
                        driver, wait, By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=2']", condition="presence"
                    )
                    self.click_element(driver, modulo_link, "link módulo faturamento")
                    time.sleep(2)
                except Exception as e:
                    log_message(f"⚠️ Falha ao selecionar módulo: {e} - acessando diretamente", "WARNING")
                    driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                    time.sleep(2)
            elif "moduloFaturamento" not in current_url:
                log_message("⚠️ URL inesperada após login - acessando módulo diretamente", "WARNING")
                driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                time.sleep(2)

            try:
                modal_close_button = driver.find_element(By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button")
                if self.headless_mode or modal_close_button.is_displayed():
                    self.click_element(driver, modal_close_button, "fechar modal inicial PathoWeb")
                    time.sleep(1)
            except Exception:
                pass

            driver.get("https://pathoweb.com.br/moduloFaturamento/index")

            log_message("Clicando em 'Preparar exames para fatura'...", "INFO")
            try:
                preparar_btn = self.wait_for_element(
                    driver, wait, By.CSS_SELECTOR,
                    "a.btn.btn-danger.chamadaAjax.setupAjax[data-url='/moduloFaturamento/preFaturamento']",
                    condition="presence"
                )
                self.click_element(driver, preparar_btn, "botão 'Preparar exames para fatura'")
            except Exception:
                preparar_btn = self.wait_for_element(
                    driver, wait, By.XPATH,
                    "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]",
                    condition="presence"
                )
                self.click_element(driver, preparar_btn, "botão 'Preparar exames para fatura' (alternativo)")

            try:
                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
            except Exception:
                time.sleep(1)

            log_message("✅ PathoWeb pronto para processamento", "SUCCESS")
            return True
        except Exception as e:
            log_message(f"❌ Erro ao fazer login no PathoWeb: {e}", "ERROR")
            return False

    def preencher_campos_exame_pathoweb(self, driver, wait, numero_guia):
        try:
            log_message("📝 Preenchendo campos do exame no PathoWeb...", "INFO")
            data_atual = datetime.now()
            ymd = data_atual.strftime("%Y-%m-%d")
            br = data_atual.strftime("%d/%m/%Y")

            time.sleep(2)

            js_data_autorizacao = f"""
            const $input = $('#requisicao_r input[name="dataAutorizacao"]').first();
            if (!$input.length) return;
            const $a = $input.closest('td').children('a.table-editable-ancora').first();
            $input.val('{ymd}').attr('value', '{ymd}')
                 .trigger('focus').trigger('input').trigger('change').trigger('blur');
            if ($a.length) $a.text('{br}').css('display', 'inline');
            """
            driver.execute_script(js_data_autorizacao)
            time.sleep(0.5)

            js_data_requisicao = f"""
            const $input = $('#requisicao_r input[name="dataRequisicao"]').first();
            if (!$input.length) return;
            const $a = $input.closest('td').children('a.table-editable-ancora').first();
            $input.val('{ymd}').attr('value', '{ymd}')
                 .trigger('focus').trigger('input').trigger('change').trigger('blur');
            if ($a.length) $a.text('{br}').css('display', 'inline');
            """
            driver.execute_script(js_data_requisicao)
            time.sleep(0.5)

            js_numero_guia = f"""
            const digitarGuia = (texto, delay = 40) => {{
              const $inp = $("#numeroGuiaInput");
              const $a   = $inp.closest('td').children('a.table-editable-ancora').first();
              $inp.val("").attr("value","").trigger("input");
              if ($a.length) $a.text("").css("display","inline");
              let i = 0;
              const timer = setInterval(() => {{
                const atual = $inp.val() + texto[i];
                $inp.val(atual).trigger("input").trigger("keyup");
                if ($a.length) $a.text(atual);
                i++;
                if (i >= texto.length) {{
                  clearInterval(timer);
                  $inp.attr("value", texto)
                      .data("previous-value", texto)
                      .trigger("change")
                      .trigger("blur");
                }}
              }}, delay);
            }};
            digitarGuia("{numero_guia}", 30);
            """
            driver.execute_script(js_numero_guia)
            time.sleep(3)

            try:
                botao_proximo = self.wait_for_element(
                    driver, wait, By.CSS_SELECTOR,
                    "a.btn.btn-sm.btn-primary.wizardControl.chamadaAjax.setupAjax[data-url='/moduloFaturamento/saveAjaxExameParaFaturamento']",
                    condition="presence"
                )
                self.click_element(driver, botao_proximo, "botão 'Próximo'")
                time.sleep(3)
            except Exception as e:
                log_message(f"⚠️ Erro ao clicar em 'Próximo': {e}", "WARNING")

            try:
                botao_salvar = self.wait_for_element(
                    driver, wait, By.CSS_SELECTOR,
                    "a.btn.btn-sm.btn-primary.chamadaAjax.setupAjax[data-url='/moduloFaturamento/saveExameDadosClinicos']",
                    condition="presence"
                )
                self.click_element(driver, botao_salvar, "botão 'Salvar'")
                time.sleep(3)
            except Exception as e:
                log_message(f"⚠️ Erro ao clicar em 'Salvar': {e}", "WARNING")

            try:
                modal = wait.until(EC.presence_of_element_located((By.ID, "myModal")))
                close_btn = modal.find_element(By.CSS_SELECTOR, "button.close[data-dismiss='modal']")
                self.click_element(driver, close_btn, "botão fechar modal PathoWeb")
                time.sleep(1)
            except Exception:
                pass

            try:
                wait.until(EC.presence_of_element_located((By.ID, "tabelaPreFaturamentoTbody")))
                self.marcar_exames_como_pendentes(driver, wait)
            except Exception as e:
                log_message(f"⚠️ Erro ao marcar exames como 'Pendente': {e}", "WARNING")

            log_message("✅ Exame atualizado no PathoWeb", "SUCCESS")
            return True
        except Exception as e:
            log_message(f"❌ Erro ao preencher dados no PathoWeb: {e}", "ERROR")
            return False

    def abrir_exame_pathoweb(self, driver, wait, numero_guia_original, numero_guia_hospital=None):
        try:
            log_message(f"🔍 Abrindo exame {numero_guia_original} no PathoWeb...", "INFO")
            campo_exame = wait.until(EC.element_to_be_clickable((By.ID, "codigoBarras")))
            campo_exame.clear()
            time.sleep(0.5)
            campo_exame.send_keys(str(numero_guia_original))
            time.sleep(0.5)

            pesquisar_btn = self.wait_for_element(driver, wait, By.ID, "pesquisaFaturamento", condition="presence")
            self.click_element(driver, pesquisar_btn, "botão Pesquisar")

            try:
                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
            except Exception:
                time.sleep(3)

            time.sleep(2)
            tbody_rows = driver.find_elements(By.CSS_SELECTOR, "#tabelaPreFaturamentoTbody tr")
            if not tbody_rows:
                log_message(f"⚠️ Nenhum resultado no PathoWeb para {numero_guia_original}", "WARNING")
                return False

            checkbox = tbody_rows[0].find_element(By.CSS_SELECTOR, "input[type='checkbox'][name='exameId']")
            if not checkbox.is_selected():
                self.click_element(driver, checkbox, "checkbox exame PathoWeb")
            time.sleep(1)

            abrir_btn = self.wait_for_element(
                driver, wait, By.CSS_SELECTOR,
                "a.btn.btn-sm.btn-primary.chamadaAjax.toogleInicial.setupAjax[data-url='/moduloFaturamento/abrirExameCorrecao']",
                condition="presence"
            )
            self.click_element(driver, abrir_btn, "botão 'Abrir exame'")
            time.sleep(2)

            try:
                modal = wait.until(EC.presence_of_element_located((By.ID, "myModal")))
                if self.headless_mode or modal.is_displayed():
                    if numero_guia_hospital:
                        self.preencher_campos_exame_pathoweb(driver, wait, numero_guia_hospital)
                    return True
            except Exception as e:
                log_message(f"⚠️ Modal do PathoWeb não encontrado: {e}", "WARNING")
                return False
        except Exception as e:
            log_message(f"❌ Erro ao abrir exame {numero_guia_original} no PathoWeb: {e}", "ERROR")
            return False

    def marcar_exames_como_pendentes(self, driver, wait):
        try:
            log_message("📝 Definindo situação como 'Pendente' na lista PathoWeb...", "INFO")
            time.sleep(1.5)

            def obter_linhas():
                return driver.find_elements(By.CSS_SELECTOR, "#tabelaPreFaturamentoTbody tr")

            linhas = obter_linhas()
            if not linhas:
                log_message("⚠️ Nenhuma linha encontrada na tabela PathoWeb", "WARNING")
                return

            total = len(linhas)
            processadas = 0

            for idx in range(total):
                log_message(f"🔄 Processando linha {idx + 1}/{total}...", "INFO")
                marcou = False
                for tentativa in range(4):
                    try:
                        try:
                            WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.ID, "spinner")))
                            WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                        except Exception:
                            pass

                        time.sleep(0.3)
                        linhas_atual = obter_linhas()
                        if idx >= len(linhas_atual):
                            raise Exception("linha não disponível")

                        linha = linhas_atual[idx]
                        celulas = linha.find_elements(By.CSS_SELECTOR, "td")
                        if len(celulas) < 2:
                            raise Exception("células insuficientes")

                        cel_status = celulas[1]
                        ancora = cel_status.find_element(By.CSS_SELECTOR, "a.table-editable-ancora")
                        texto_atual = (ancora.text or "").strip().lower()
                        if texto_atual == "pendente":
                            log_message(f"✅ Linha {idx + 1}: já está 'Pendente'", "SUCCESS")
                            processadas += 1
                            break

                        self.click_element(driver, ancora, f"âncora status linha {idx + 1}")
                        time.sleep(0.3)

                        linhas_temp = obter_linhas()
                        if idx >= len(linhas_temp):
                            raise Exception("linha não disponível após clique")
                        cel_status_temp = linhas_temp[idx].find_elements(By.CSS_SELECTOR, "td")[1]
                        select_el = cel_status_temp.find_element(By.CSS_SELECTOR, "select[name='faturamentoConferido']")
                        driver.execute_script(
                            """
                            var s = arguments[0];
                            $(s).val('Pendente').trigger('change').trigger('blur');
                            """,
                            select_el
                        )

                        try:
                            WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "spinner")))
                            WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                        except Exception:
                            time.sleep(0.3)

                        processadas += 1
                        log_message(f"✅ Linha {idx + 1}: marcada como 'Pendente'", "SUCCESS")
                        marcou = True
                        break

                    except StaleElementReferenceException:
                        log_message(f"⚠️ Linha {idx + 1}: elemento inválido após atualização (tentativa {tentativa + 1})", "WARNING")
                        time.sleep(0.3)
                        continue
                    except Exception as e:
                        log_message(f"⚠️ Linha {idx + 1}: tentativa {tentativa + 1} falhou: {e}", "WARNING")
                        time.sleep(0.3)
                        continue
                if not marcou:
                    log_message(f"❌ Não foi possível marcar linha {idx + 1} como 'Pendente' após tentativas", "ERROR")

            log_message(f"📊 Linhas marcadas como 'Pendente': {processadas}/{total}", "INFO")
        except Exception as e:
            log_message(f"❌ Erro ao marcar 'Pendente' na tabela PathoWeb: {e}", "ERROR")

    def autorizar_guia(self, driver, wait):
        log_message("🔄 Clicando em Autorizar...", "INFO")
        botao_autorizar = self.wait_for_element(driver, wait, By.ID, "autorizar", condition="clickable")
        try:
            botao_autorizar.click()
        except Exception:
            self.click_element(driver, botao_autorizar, "botão Autorizar (fallback)")

        time.sleep(3)

        indicadores_sucesso = driver.find_elements(By.CSS_SELECTOR, ".alert-success, .alert.alert-success")
        for alerta in indicadores_sucesso:
            if alerta.is_displayed():
                texto = alerta.text.strip()
                numero = self._extrair_numero_guia(texto)
                log_message(f"✅ Guia autorizada: {texto}", "SUCCESS")
                return {'sucesso': True, 'numero_guia': numero, 'mensagem': texto}

        indicadores_info = driver.find_elements(By.CSS_SELECTOR, ".alert-info, .alert.alert-info")
        for alerta in indicadores_info:
            if alerta.is_displayed():
                texto = alerta.text.strip()
                numero = self._extrair_numero_guia(texto)
                log_message(f"ℹ️ Mensagem após autorização: {texto}", "INFO")
                return {'sucesso': True, 'numero_guia': numero, 'mensagem': texto}

        indicadores_erro = driver.find_elements(By.CSS_SELECTOR, ".alert-danger, .alert.alert-danger, #msg_erro")
        for alerta in indicadores_erro:
            if alerta.is_displayed():
                texto = alerta.text.strip()
                if not texto:
                    texto = alerta.get_attribute("innerText").strip()
                log_message(f"❌ Erro na autorização: {texto}", "ERROR")
                return {'sucesso': False, 'mensagem': texto}

        modal = driver.find_elements(By.CSS_SELECTOR, ".modal.in, .modal.show")
        for elemento in modal:
            if elemento.is_displayed():
                texto = elemento.text.strip()
                numero = self._extrair_numero_guia(texto)
                log_message(f"ℹ️ Modal após autorização: {texto}", "INFO")
                return {'sucesso': True, 'numero_guia': numero, 'mensagem': texto}

        log_message("⚠️ Não foi possível confirmar o resultado da autorização", "WARNING")
        return {'sucesso': False, 'mensagem': 'Não foi possível confirmar o resultado da autorização'}

    def _extrair_numero_guia(self, texto):
        if not texto:
            return ''
        match = re.search(r'(\d{6,})', texto)
        return match.group(1) if match else ''

    def limpar_mensagem_erro(self, mensagem):
        if not mensagem:
            return ''
        msg_str = str(mensagem)
        if 'Stacktrace:' in msg_str or 'Session info:' in msg_str:
            linhas = msg_str.split('\n')
            for linha in linhas:
                if linha.strip() and not linha.strip().startswith('('):
                    return linha.strip()[:200]
        return msg_str[:200]

    def processar_guia(self, driver, wait, dados):
        log_message(f"🔄 Processando guia {dados['guia']}", "INFO")

        if not dados['cartao']:
            log_message("⚠️ Cartão vazio, guia pulada", "WARNING")
            return {
                'guia': dados['guia'],
                'status': 'erro',
                'erro': 'Cartão vazio',
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

        resultado_cartao = self.preencher_codigo_beneficiario(driver, wait, dados['cartao'])
        if resultado_cartao.get('erro'):
            return {
                'guia': dados['guia'],
                'status': 'erro',
                'erro': resultado_cartao.get('mensagem'),
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

        try:
            self.buscar_medico_solicitante(driver, wait, dados['medico'], dados.get('crm'))
        except Exception as e:
            log_message(f"⚠️ Erro ao buscar médico: {e}", "WARNING")
            return {
                'guia': dados['guia'],
                'status': 'erro',
                'erro': f"Erro ao buscar médico: {e}",
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

        self.preencher_campos_fixos(driver, wait)

        try:
            self.preencher_procedimentos(driver, dados['procedimentos'], dados['qtd'])
        except Exception as e:
            log_message(f"⚠️ Erro ao preencher procedimentos: {e}", "WARNING")
            return {
                'guia': dados['guia'],
                'status': 'erro',
                'erro': f"Erro ao preencher procedimentos: {e}",
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

        resultado_autorizacao = self.autorizar_guia(driver, wait)
        if resultado_autorizacao.get('sucesso'):
            numero_guia = resultado_autorizacao.get('numero_guia')
            if numero_guia:
                log_message(f"✅ Guia {dados['guia']} autorizada - Nº {numero_guia}", "SUCCESS")
            return {
                'guia': dados['guia'],
                'status': 'sucesso',
                'numero_guia': numero_guia,
                'mensagem': resultado_autorizacao.get('mensagem'),
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

        return {
            'guia': dados['guia'],
            'status': 'erro',
            'erro': resultado_autorizacao.get('mensagem'),
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

    def salvar_resultados_excel(self, excel_file, resultados):
        try:
            log_message("💾 Salvando resultados no Excel...", "INFO")
            df = pd.read_excel(excel_file, header=0, dtype={'CARTAO': str})
            df.columns = [col.strip() for col in df.columns]

            colunas_adicionadas = []
            if 'Numero_Guia' not in df.columns and 'NUMERO_GUIA' not in df.columns:
                df['Numero_Guia'] = ''
                colunas_adicionadas.append('Numero_Guia')
            if 'Status_Processamento' not in df.columns and 'STATUS_PROCESSAMENTO' not in df.columns:
                df['Status_Processamento'] = ''
                colunas_adicionadas.append('Status_Processamento')
            if 'Status_Guia' not in df.columns and 'STATUS_GUIA' not in df.columns:
                df['Status_Guia'] = ''
                colunas_adicionadas.append('Status_Guia')
            if 'Mensagem_Erro' not in df.columns and 'MENSAGEM_ERRO' not in df.columns:
                df['Mensagem_Erro'] = ''
                colunas_adicionadas.append('Mensagem_Erro')
            if 'Data_Processamento' not in df.columns and 'DATA_PROCESSAMENTO' not in df.columns:
                df['Data_Processamento'] = ''
                colunas_adicionadas.append('Data_Processamento')

            if colunas_adicionadas:
                log_message(f"✅ Colunas adicionadas: {colunas_adicionadas}", "SUCCESS")

            def localizar_coluna(nome):
                for col in df.columns:
                    if col.upper().strip() == nome.upper().strip():
                        return col
                return None

            coluna_guia = localizar_coluna('GUIA')
            if not coluna_guia:
                raise ValueError("Coluna GUIA não encontrada no Excel")

            col_numero_guia = localizar_coluna('NUMERO_GUIA') or 'Numero_Guia'
            col_status_proc = localizar_coluna('STATUS_PROCESSAMENTO') or 'Status_Processamento'
            col_status_guia = localizar_coluna('STATUS_GUIA') or 'Status_Guia'
            col_mensagem_erro = localizar_coluna('MENSAGEM_ERRO') or 'Mensagem_Erro'
            col_data_proc = localizar_coluna('DATA_PROCESSAMENTO') or 'Data_Processamento'

            for resultado in resultados:
                guia = resultado.get('guia')
                mask = df[coluna_guia].astype(str).str.strip() == str(guia).strip()
                indices = df[mask].index
                if len(indices) == 0:
                    log_message(f"⚠️ Guia {guia} não encontrada no Excel", "WARNING")
                    continue

                idx = indices[0]
                df.loc[idx, col_data_proc] = resultado.get('timestamp', '')
                df.loc[idx, col_status_guia] = resultado.get('status_guia', '')
                df.loc[idx, col_numero_guia] = resultado.get('numero_guia', '')

                status = resultado.get('status')
                if status == 'sucesso':
                    df.loc[idx, col_status_proc] = 'SUCESSO'
                    df.loc[idx, col_mensagem_erro] = ''
                else:
                    df.loc[idx, col_status_proc] = 'ERRO'
                    mensagem = resultado.get('erro', resultado.get('mensagem', ''))
                    df.loc[idx, col_mensagem_erro] = self.limpar_mensagem_erro(mensagem)

            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')

            log_message(f"✅ Resultados salvos em {excel_file}", "SUCCESS")
            return excel_file

        except Exception as e:
            log_message(f"❌ Erro ao salvar resultados: {e}", "ERROR")
            return None

    def run(self, params: dict):
        username = params.get("hospital_user")
        password = params.get("hospital_pass")
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode")
        excel_file = params.get("excel_file")

        self.headless_mode = headless_mode
        log_message(f"🔧 Modo headless: {'Ativado' if headless_mode else 'Desativado'}", "INFO")

        if not username or not password:
            messagebox.showerror("Erro", "Credenciais do Hospitalar são obrigatórias.")
            return

        if not excel_file or not os.path.exists(excel_file):
            messagebox.showerror("Erro", "Arquivo Excel é obrigatório para este módulo.")
            return

        driver = BrowserFactory.create_chrome(headless=headless_mode)
        wait = WebDriverWait(driver, 15)

        try:
            log_message("Iniciando automação de Lançamento de Guia Hospitalar...", "INFO")

            try:
                dados_excel = self.read_excel_data(excel_file)
                log_message(f"✅ Carregados {len(dados_excel)} registros do Excel", "SUCCESS")
            except Exception as e:
                log_message(f"❌ Erro ao ler Excel: {e}", "ERROR")
                messagebox.showerror("Erro", f"Erro ao ler arquivo Excel:\n{e}")
                return

            if not dados_excel:
                log_message("ℹ️ Nenhum registro para processar", "INFO")
                messagebox.showinfo("Informação", "Nenhum registro encontrado no Excel.")
                return

            self.fazer_login(driver, wait, username, password)
            self.navegar_para_guia_procedimento(driver, wait)

            resultados = []

            for indice, dados in enumerate(dados_excel, start=1):
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break

                log_message(f"➡️ Processando {indice}/{len(dados_excel)} - Guia {dados['guia']}", "INFO")
                try:
                    resultado = self.processar_guia(driver, wait, dados)
                    resultados.append(resultado)
                except Exception as e:
                    log_message(f"❌ Erro inesperado na guia {dados['guia']}: {e}", "ERROR")
                    resultados.append({
                        'guia': dados['guia'],
                        'status': 'erro',
                        'erro': str(e),
                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })

                if indice < len(dados_excel):
                    self.navegar_para_guia_procedimento(driver, wait)

            pathoweb_user = params.get("username")
            pathoweb_pass = params.get("password")
            pathoweb_sucessos = 0
            guias_para_pathoweb = [
                r for r in resultados if r.get('status') == 'sucesso' and r.get('numero_guia')
            ]

            if guias_para_pathoweb:
                if pathoweb_user and pathoweb_pass:
                    log_message("🌐 Iniciando atualização das guias no PathoWeb...", "INFO")
                    if self.fazer_login_pathoweb(driver, wait, pathoweb_user, pathoweb_pass):
                        for idx, resultado in enumerate(guias_para_pathoweb, 1):
                            if cancel_flag and cancel_flag.is_set():
                                log_message("Execução cancelada pelo usuário durante atualização no PathoWeb.", "WARNING")
                                break

                            numero_original = resultado.get('guia')
                            numero_guia_portal = resultado.get('numero_guia')
                            log_message(f"🔁 Atualizando exame {idx}/{len(guias_para_pathoweb)} "
                                        f"(Guia original {numero_original}, nº portal {numero_guia_portal})", "INFO")

                            try:
                                sucesso = self.abrir_exame_pathoweb(driver, wait, numero_original, numero_guia_portal)
                                if sucesso:
                                    pathoweb_sucessos += 1
                                    log_message("✅ Atualização no PathoWeb concluída para este exame", "SUCCESS")
                                else:
                                    log_message("⚠️ Não foi possível atualizar este exame no PathoWeb", "WARNING")
                            except Exception as e:
                                log_message(f"❌ Erro ao atualizar exame {numero_original} no PathoWeb: {e}", "ERROR")

                            if idx < len(guias_para_pathoweb):
                                try:
                                    driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                                    time.sleep(2)
                                    preparar_btn = self.wait_for_element(
                                        driver, wait, By.CSS_SELECTOR,
                                        "a.btn.btn-danger.chamadaAjax.setupAjax[data-url='/moduloFaturamento/preFaturamento']",
                                        condition="presence"
                                    )
                                    self.click_element(driver, preparar_btn, "botão 'Preparar exames para fatura' (reload)")
                                    try:
                                        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                                        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                                    except Exception:
                                        time.sleep(1)
                                except Exception as e:
                                    log_message(f"⚠️ Erro ao recarregar tela do PathoWeb: {e}", "WARNING")
                    else:
                        log_message("❌ Falha no login do PathoWeb - guias não foram atualizadas.", "ERROR")
                else:
                    log_message("ℹ️ Credenciais PathoWeb não informadas - pulando atualização no PathoWeb.", "INFO")
            else:
                log_message("ℹ️ Nenhuma guia aprovada para atualizar no PathoWeb.", "INFO")

            arquivo_resultados = self.salvar_resultados_excel(excel_file, resultados)

            total = len(resultados)
            sucessos = sum(1 for r in resultados if r.get('status') == 'sucesso')
            erros = sum(1 for r in resultados if r.get('status') == 'erro')

            log_message("\n📊 Resumo do processamento Hospitalar:", "INFO")
            log_message(f"Total de registros: {total}", "INFO")
            log_message(f"Sucessos: {sucessos}", "SUCCESS" if sucessos else "INFO")
            log_message(f"Erros: {erros}", "ERROR" if erros else "INFO")
            log_message(f"Atualizações PathoWeb: {pathoweb_sucessos}", "INFO")

            mensagem_final = (
                "✅ Processamento concluído!\n\n"
                f"Total de registros: {total}\n"
                f"Sucessos: {sucessos}\n"
                f"Erros: {erros}\n"
                f"Atualizações PathoWeb: {pathoweb_sucessos}"
            )

            if arquivo_resultados:
                mensagem_final += f"\n\n📊 Resultados salvos em:\n{arquivo_resultados}"

            messagebox.showinfo("Processamento Concluído", mensagem_final)

            return {
                'sucesso': sucessos > 0,
                'pathoweb_sucessos': pathoweb_sucessos,
                'resultados': resultados,
                'arquivo_resultados': arquivo_resultados
            }

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            if not headless_mode:
                try:
                    input("Pressione Enter para fechar o navegador...")
                except Exception:
                    pass
            driver.quit()


def run(params: dict):
    module = LancamentoGuiaHospitalarModule()
    module.run(params)
