import os
import re
import pandas as pd
import numpy as np
from time import sleep
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    TimeoutException,
    NoSuchElementException,
)

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule


class LancamentoFinanceiroModule(BaseModule):
    LOGIN_URL = "https://dap.pathoweb.com.br/login/auth?format="

    def __init__(self):
        super().__init__("Lançamento Financeiro")
        self.driver = None
        self.wait = None
        self.wait_fast = None
        self.headless_mode = False
        self.cancel_flag = None
        self.conta_atual = None

    @staticmethod
    def clean_numero_documento(value, keep_leading_zeros=False):
        if pd.isna(value):
            return ""
        text = str(value).strip()
        digits = re.sub(r"\D", "", text)
        if not digits:
            return ""
        if not keep_leading_zeros:
            digits = digits.lstrip("0")
            if digits == "":
                digits = "0"
        return digits

    @staticmethod
    def clean_parcelas(value):
        if pd.isna(value):
            return ""
        text = str(value).strip()
        normalized = text.replace(",", ".")
        try:
            number = float(normalized)
            integer = int(number)
            if abs(number - integer) < 1e-8:
                return str(integer)
            return str(number).rstrip("0").rstrip(".")
        except Exception:
            digits = re.sub(r"\D", "", text)
            return digits.lstrip("0")

    MONTH_MAP = {
        "jan": "Jan",
        "fev": "Feb",
        "mar": "Mar",
        "abr": "Apr",
        "mai": "May",
        "jun": "Jun",
        "jul": "Jul",
        "ago": "Aug",
        "set": "Sep",
        "out": "Oct",
        "nov": "Nov",
        "dez": "Dec",
    }

    @classmethod
    def _normalize_date_text(cls, value: str) -> str:
        normalized = value.lower()
        for pt, en in cls.MONTH_MAP.items():
            normalized = re.sub(pt, en.lower(), normalized)
        return normalized

    @classmethod
    def format_datetime(cls, value, fmt, min_year=1900, max_year=2100):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
        
        # Se for número (serial do Excel), converter
        if isinstance(value, (int, float)) and not pd.isna(value):
            try:
                # Excel serial date: 1 = 1900-01-01
                if 1 <= value <= 1000000:  # Range razoável para datas Excel
                    from datetime import datetime, timedelta
                    excel_epoch = datetime(1899, 12, 30)
                    parsed = excel_epoch + timedelta(days=value)
                    if parsed.year >= min_year and parsed.year <= max_year:
                        return parsed.strftime(fmt)
                    else:
                        log_message(f"⚠️ Data Excel fora do intervalo: {value} → {parsed.strftime('%Y-%m-%d')}", "WARNING")
                        return ""
            except Exception:
                pass
        
        # Se for string, processar
        if isinstance(value, str):
            value = value.strip()
            if not value:
                return ""
            original_value = value
            value = cls._normalize_date_text(value)
        else:
            original_value = str(value)
        
        # Tentar múltiplos formatos
        parsed = None
        formatos = [
            ("%d/%m/%Y", False),  # DD/MM/YYYY
            ("%d-%m-%Y", False),  # DD-MM-YYYY
            ("%Y-%m-%d", False),  # YYYY-MM-DD
            ("%d/%m/%y", False),  # DD/MM/YY (2 dígitos)
            ("%m/%d/%Y", False),  # MM/DD/YYYY (formato americano)
            (None, True),  # pd.to_datetime com dayfirst=True
        ]
        
        for formato, usar_pandas in formatos:
            try:
                if usar_pandas:
                    parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
                else:
                    parsed = pd.to_datetime(value, format=formato, errors="coerce")
                
                if pd.notna(parsed):
                    # Validar ano
                    if parsed.year < min_year or parsed.year > max_year:
                        log_message(
                            f"⚠️ Data fora do intervalo ({original_value} → {parsed.strftime('%Y-%m-%d')}), ignorando",
                            "WARNING"
                        )
                        return ""
                    return parsed.strftime(fmt)
            except Exception:
                continue
        
        log_message(f"⚠️ Data inválida ignorada: {original_value}", "WARNING")
        return ""

    def prepare_dataframe(self, excel_file):
        try:
            df = pd.read_excel(excel_file)
        except Exception as exc:
            raise ValueError(f"Erro ao ler o Excel: {exc}") from exc

        df = df.fillna("")
        if "Valor da transação" not in df.columns:
            raise ValueError("Planilha precisa da coluna 'Valor da transação'.")

        df["Valor da transação"] = df["Valor da transação"].replace("", np.nan)
        df["Valor da transação"] = pd.to_numeric(df["Valor da transação"], errors="coerce")
        df["Valor da transação"] = df["Valor da transação"].fillna(0).astype(float)
        df["Valor da transação"] = (
            df["Valor da transação"].map("{:.2f}".format).str.replace(".", ",")
        )
        return df

    def setup_browser(self, headless):
        self.headless_mode = headless
        self.driver = BrowserFactory.create_chrome(headless=headless)
        self.wait = WebDriverWait(self.driver, 12)
        self.wait_fast = WebDriverWait(self.driver, 4)

    def close_browser(self):
        if self.driver:
            log_message("🔒 Encerrando navegador...", "INFO")
            self.driver.quit()
            self.driver = None

    def click_element(self, element, descricao="elemento"):
        try:
            element.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", element)
        except Exception:
            self.driver.execute_script("arguments[0].click();", element)
        log_message(f"✅ Clique em {descricao}", "INFO")

    def realizar_login(self, username, password):
        log_message("🔐 Acessando página de login...", "INFO")
        self.driver.get(self.LOGIN_URL)
        sleep(2)

        email_field = None
        try:
            email_field = self.wait.until(EC.element_to_be_clickable((By.NAME, "email")))
        except TimeoutException:
            email_field = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='email']"))
            )

        password_field = None
        try:
            password_field = self.driver.find_element(By.NAME, "password")
        except NoSuchElementException:
            password_field = self.driver.find_element(By.CSS_SELECTOR, "input[type='password']")

        email_field.clear()
        email_field.send_keys(username)
        password_field.clear()
        password_field.send_keys(password)

        btn_login = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' or contains(., 'Entrar')]"))
        )
        self.click_element(btn_login, "Entrar")
        sleep(3)

    def acessar_menu_financeiro(self):
        log_message("🧭 Abrindo módulo Financeiro...", "INFO")
        menu = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'Financeiro')]")))
        self.click_element(menu, "Menu Financeiro")
        sleep(2)

    def fechar_modal_inicial(self):
        seletores = [
            "//div[@id='mensagemParaClienteModal']//button[@data-dismiss='modal' or contains(., 'Fechar')]",
            "//button[contains(text(), 'Fechar') and contains(@class, 'btn-default')]",
            "//div[contains(@class, 'modal')]//button[contains(text(), 'Fechar')]",
        ]
        for xpath in seletores:
            try:
                fechar = self.wait_fast.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                self.click_element(fechar, "Modal inicial")
                sleep(0.4)
                break
            except TimeoutException:
                continue

    def selecionar_conta(self, conta_nome):
        if not conta_nome:
            return False
        try:
            self.driver.execute_script("$('#bancoSelect').select2('open');")
            sleep(0.4)
            search = self.wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, ".select2-search__field"))
            )
            search.send_keys(conta_nome)
            sleep(0.7)
            primeira_opcao = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-results__option"))
            )
            self.click_element(primeira_opcao, f"Conta {conta_nome}")
            self.conta_atual = conta_nome
            sleep(0.4)
            return True
        except Exception as exc:
            log_message(f"⚠️ Erro ao selecionar conta '{conta_nome}': {exc}", "WARNING")
            return False

    def selecionar_primeira_conta(self, df):
        for _, row in df.iterrows():
            conta = str(row.get("Conta", "")).strip()
            if conta:
                if self.selecionar_conta(conta):
                    log_message(f"✅ Conta inicial selecionada: {conta}", "SUCCESS")
                else:
                    log_message("❌ Não foi possível selecionar a conta inicial", "ERROR")
                break

    def garantir_conta(self, conta_desejada):
        conta_desejada = str(conta_desejada).strip()
        if not conta_desejada:
            raise ValueError("Conta não informada na planilha.")
        if self.conta_atual == conta_desejada:
            return
        if not self.selecionar_conta(conta_desejada):
            raise ValueError(f"Não foi possível trocar para a conta {conta_desejada}.")

    def abrir_formulario(self, tipo_lancamento):
        if tipo_lancamento == "receita":
            botao_xpath = "//div[@id='buttonsActionsFinan']//a[contains(@class, 'btn-primary') and contains(., 'Adicionar')]"
            data_url = "/moduloFinanceiro/adicionarTransacaoDinheiro/0"
        else:
            botao_xpath = "//div[@id='buttonsActionsFinan']//a[contains(@class, 'btn-danger') and contains(., 'Adicionar')]"
            data_url = "/moduloFinanceiro/adicionarTransacaoDinheiro/1"

        botao = self.wait.until(EC.element_to_be_clickable((By.XPATH, botao_xpath)))
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)
        self.click_element(botao, f"Abrir menu {tipo_lancamento}")
        sleep(0.3)

        dinheiro_xpath = f"//ul[contains(@class, 'dropdown-menu')]//a[contains(@data-url, '{data_url}')]"
        dinheiro = self.wait.until(EC.element_to_be_clickable((By.XPATH, dinheiro_xpath)))
        self.click_element(dinheiro, f"Dinheiro ({tipo_lancamento})")
        sleep(1.2)

    def selecionar_option_por_texto(self, select_element, valor, descricao):
        if not valor:
            return
        if isinstance(select_element, Select):
            select = select_element
        else:
            select = Select(select_element)
        texto_limpo = valor.strip()
        try:
            select.select_by_visible_text(texto_limpo)
            return
        except Exception:
            for option in select.options:
                if option.text.strip().lower() == texto_limpo.lower():
                    option.click()
                    return
        log_message(f"⚠️ Não foi possível selecionar '{valor}' em {descricao}", "WARNING")

    def preencher_conta_form(self, conta_nome):
        if not conta_nome:
            return
        try:
            seletor = self.wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "#formaPagamentoForm #bancoSelect"))
            )
        except TimeoutException:
            elementos = self.driver.find_elements(By.ID, "bancoSelect")
            if not elementos:
                log_message("⚠️ Campo de conta não encontrado no formulário", "WARNING")
                return
            seletor = elementos[-1]
        try:
            self.selecionar_option_por_texto(seletor, conta_nome, "Conta do lançamento")
        except Exception as exc:
            log_message(f"⚠️ Não foi possível definir a conta no formulário: {exc}", "WARNING")

    def preencher_envolvido(self, envolvido):
        if not envolvido:
            return
        try:
            self.driver.execute_script("$('#envolvido').select2('open');")
            sleep(0.4)
            search = self.wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, ".select2-search__field"))
            )
            search.send_keys(envolvido)
            sleep(0.8)
            primeira = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-results__option"))
            )
            self.click_element(primeira, f"Envolvido {envolvido}")
        except Exception:
            try:
                campo = self.wait.until(EC.visibility_of_element_located((By.NAME, "envolvido")))
                campo.clear()
                campo.send_keys(envolvido)
            except Exception as exc:
                log_message(f"⚠️ Não foi possível preencher Envolvido: {exc}", "WARNING")

    def preencher_formulario(self, row):
        numero_documento = self.clean_numero_documento(row.get("Número do documento"))
        numero_parcelas = self.clean_parcelas(row.get("Número de Parcelas"))
        valor_transacao = str(row.get("Valor da transação", "0,00")).strip()
        competencia = self.format_datetime(row.get("Data de Competência"), "%Y-%m")
        data_prevista = self.format_datetime(row.get("Data prevista"), "%Y-%m-%d")
        data_lancamento = self.format_datetime(row.get("Data de Lançamento"), "%Y-%m-%d")
        # Para datetime-local, usar formato YYYY-MM-DDTHH:mm
        data_efetivada_raw = self.format_datetime(row.get("Data efetivada"), "%Y-%m-%d")
        data_efetivada = f"{data_efetivada_raw}T00:00" if data_efetivada_raw else ""
        observacao = str(row.get("Observação", "")).strip()
        tipo_transacao = str(row.get("Tipo de transação", "")).strip()
        tipo = str(row.get("Tipo", "")).strip()
        envolvido = str(row.get("Envolvido", "")).strip()
        # Dias para vencimento (Coluna L)
        dias_vencimento = str(row.get("Dias para Vencimento", "")).strip()
        if not dias_vencimento:
            dias_vencimento = "30"
        try:
            dias_int = int(float(dias_vencimento))
            dias_vencimento = str(max(dias_int, 0))
        except (ValueError, TypeError):
            dias_vencimento = "30"

        # Preencher datas via JavaScript
        if competencia:
            self.driver.execute_script(f"document.querySelector('#dataCompetencia').value = '{competencia}';")

        if numero_documento:
            campo = self.wait.until(EC.visibility_of_element_located((By.ID, "numeroDocumento")))
            campo.clear()
            campo.send_keys(numero_documento)

        campo_valor = self.wait.until(EC.element_to_be_clickable((By.ID, "valorTransacao")))
        campo_valor.click()
        campo_valor.send_keys(Keys.CONTROL + "a")
        campo_valor.send_keys(Keys.DELETE)
        sleep(0.2)
        campo_valor.send_keys(valor_transacao)

        try:
            select_tipo_transacao = Select(
                self.wait.until(EC.visibility_of_element_located((By.ID, "tipoTransacao.id")))
            )
            self.selecionar_option_por_texto(select_tipo_transacao, tipo_transacao, "Tipo de transação")
        except Exception:
            pass

        self.preencher_envolvido(envolvido)

        try:
            select_tipo = Select(self.wait.until(EC.visibility_of_element_located((By.ID, "tipo"))))
            self.selecionar_option_por_texto(select_tipo, tipo, "Tipo")
        except Exception:
            pass

        if numero_parcelas:
            campo = self.wait.until(EC.visibility_of_element_located((By.ID, "numeroParcelas")))
            campo.clear()
            campo.send_keys(numero_parcelas)

        # Preencher datas via JavaScript
        if data_lancamento:
            self.driver.execute_script(f"document.querySelector('#dataLancamento').value = '{data_lancamento}';")

        if data_prevista:
            self.driver.execute_script(f"document.querySelector('#dataPrevista').value = '{data_prevista}';")

        if data_efetivada:
            # Para datetime-local, tentar por ID primeiro, depois por NAME
            try:
                self.driver.execute_script(f"document.querySelector('#dataEfetivada').value = '{data_efetivada}';")
            except Exception:
                try:
                    self.driver.execute_script(f"document.querySelector('input[name=\"dataEfetivada\"]').value = '{data_efetivada}';")
                except Exception:
                    log_message(f"⚠️ Não foi possível preencher dataEfetivada via JS", "WARNING")

        campo_dias = self.wait.until(EC.visibility_of_element_located((By.NAME, "numeroDias")))
        campo_dias.clear()
        campo_dias.send_keys(str(dias_vencimento))

        if observacao:
            campo = self.wait.until(EC.visibility_of_element_located((By.ID, "observacao")))
            campo.clear()
            campo.send_keys(observacao)

    def salvar_transacao(self, tipo_lancamento):
        btn_salvar = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, '//a[contains(@class, "btn") and contains(., "Salvar")]'))
        )
        self.click_element(btn_salvar, "Salvar transação")
        sleep(1.5)

        texto_retorno = "Adicionar receita" if tipo_lancamento == "receita" else "Adicionar despesa"
        self.wait.until(EC.visibility_of_element_located((By.XPATH, f"//*[contains(., '{texto_retorno}')]")))

    def processar_linha(self, indice, row):
        conta_atual = str(row.get("Conta", "")).strip()
        if not conta_atual:
            log_message(f"⚠️ Linha {indice + 1} ignorada (conta vazia)", "WARNING")
            return {"linha": indice + 1, "status": "ignorada", "motivo": "conta vazia"}

        tipo_lancamento = str(row.get("Lançamento", "")).strip().lower()
        if tipo_lancamento not in {"receita", "despesa"}:
            log_message(f"⚠️ Linha {indice + 1} ignorada (tipo inválido)", "WARNING")
            return {"linha": indice + 1, "status": "ignorada", "motivo": "tipo inválido"}

        try:
            self.garantir_conta(conta_atual)
            self.abrir_formulario(tipo_lancamento)
            self.preencher_conta_form(conta_atual)
            self.preencher_formulario(row)
            self.salvar_transacao(tipo_lancamento)
            log_message(f"✅ Linha {indice + 1} lançada com sucesso", "SUCCESS")
            return {"linha": indice + 1, "status": "ok"}
        except Exception as exc:
            log_message(f"❌ Erro linha {indice + 1}: {exc}", "ERROR")
            return {"linha": indice + 1, "status": "erro", "erro": str(exc)}

    def gerar_resumo(self, resultados):
        total = len(resultados)
        sucesso = len([r for r in resultados if r["status"] == "ok"])
        ignoradas = len([r for r in resultados if r["status"] == "ignorada"])
        erros = [r for r in resultados if r["status"] == "erro"]

        log_message("📊 Resumo do processamento", "INFO")
        log_message(f"Total: {total}", "INFO")
        log_message(f"✅ Sucesso: {sucesso}", "SUCCESS")
        log_message(f"⚠️ Ignoradas: {ignoradas}", "WARNING")
        log_message(f"❌ Erros: {len(erros)}", "ERROR")

        if erros:
            detalhes = "\n".join([f"Linha {r['linha']}: {r['erro']}" for r in erros[:5]])
        else:
            detalhes = "Nenhum erro reportado."

        messagebox.showinfo(
            "Resumo do Lançamento",
            f"Total processado: {total}\n"
            f"Sucesso: {sucesso}\n"
            f"Ignoradas: {ignoradas}\n"
            f"Erros: {len(erros)}\n\n"
            f"{detalhes}",
        )

    def run(self, params: dict):
        username = params.get("username", "").strip()
        password = params.get("password", "").strip()
        excel_file = params.get("excel_file")
        headless_mode = params.get("headless_mode", True)
        self.cancel_flag = params.get("cancel_flag")

        if not username or not password:
            messagebox.showwarning("Credenciais", "Informe usuário e senha válidos.")
            return

        if not excel_file or not os.path.exists(excel_file):
            messagebox.showerror("Arquivo", "Arquivo Excel não informado ou inexistente.")
            return

        try:
            df = self.prepare_dataframe(excel_file)
        except ValueError as exc:
            messagebox.showerror("Planilha inválida", str(exc))
            return

        if df.empty:
            messagebox.showwarning("Planilha vazia", "Nenhum dado encontrado no Excel.")
            return

        resultados = []

        try:
            self.setup_browser(headless_mode)
            log_message("🚀 Iniciando lançamento financeiro...", "INFO")
            self.realizar_login(username, password)
            self.acessar_menu_financeiro()
            self.fechar_modal_inicial()
            self.selecionar_primeira_conta(df)

            for indice, row in df.iterrows():
                if self.cancel_flag and self.cancel_flag.is_set():
                    log_message("⏹️ Execução cancelada pelo usuário", "WARNING")
                    break
                resultado = self.processar_linha(indice, row)
                resultados.append(resultado)

            self.gerar_resumo(resultados)
        except Exception as exc:
            log_message(f"❌ Erro durante a automação: {exc}", "ERROR")
            messagebox.showerror("Erro", f"Ocorreu um erro durante o processo:\n{exc}")
        finally:
            self.close_browser()


def run(params: dict):
    module = LancamentoFinanceiroModule()
    module.run(params)