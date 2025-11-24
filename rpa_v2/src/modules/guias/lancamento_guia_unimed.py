import os
import time
import pandas as pd
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from datetime import datetime
import re
import difflib

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

class LancamentoGuiaUnimedModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Lançamento Guia Unimed")
        self.headless_mode = False  # Será definido no run()

    def click_element(self, driver, element, descricao="elemento"):
        """Clica em um elemento de forma robusta, funcionando em modo headless e normal"""
        try:
            if self.headless_mode:
                # Em modo headless, usar sempre JavaScript para cliques mais confiáveis
                driver.execute_script("arguments[0].click();", element)
                log_message(f"✅ Clique via JavaScript em {descricao}", "INFO")
            else:
                # Em modo normal, tentar clique normal primeiro
                try:
                    element.click()
                    log_message(f"✅ Clique normal em {descricao}", "INFO")
                except Exception:
                    # Se falhar, usar JavaScript como fallback
                    driver.execute_script("arguments[0].click();", element)
                    log_message(f"✅ Clique via JavaScript (fallback) em {descricao}", "INFO")
        except Exception as e:
            log_message(f"❌ Erro ao clicar em {descricao}: {e}", "ERROR")
            raise

    def wait_for_element(self, driver, wait, by, value, condition="presence", timeout=None):
        """Aguarda elemento de forma compatível com headless"""
        try:
            if timeout:
                wait = WebDriverWait(driver, timeout)
            
            # Em modo headless, sempre usar 'presence' em vez de 'clickable' ou 'visible'
            if self.headless_mode and condition in ["clickable", "visible"]:
                element = wait.until(EC.presence_of_element_located((by, value)))
            elif condition == "clickable":
                element = wait.until(EC.element_to_be_clickable((by, value)))
            elif condition == "visible":
                element = wait.until(EC.visibility_of_element_located((by, value)))
            else:  # presence
                element = wait.until(EC.presence_of_element_located((by, value)))
            
            return element
        except Exception as e:
            log_message(f"❌ Erro ao aguardar elemento {value}: {e}", "ERROR")
            raise

    def read_excel_data(self, file_path: str) -> list:
        """Lê os dados do arquivo Excel com a estrutura: GUIA, CARTAO, MEDICO, CRM, PROCEDIMENTOS, QTD, TEXTO"""
        try:
            df = pd.read_excel(file_path, header=0)
            
            # Verificar se as colunas estão corretas
            expected_columns = ['GUIA', 'CARTAO', 'MEDICO', 'CRM', 'PROCEDIMENTOS', 'QTD', 'TEXTO']
            
            # Ajustar nomes das colunas se necessário (case insensitive)
            df.columns = df.columns.str.upper().str.strip()
            
            # Verificar se todas as colunas necessárias existem
            missing_columns = [col for col in expected_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Colunas faltando no Excel: {missing_columns}")
            
            # Converter DataFrame para lista de dicionários, removendo linhas vazias
            data_list = []
            for _, row in df.iterrows():
                if pd.notna(row['GUIA']) and str(row['GUIA']).strip():
                    # Função auxiliar para converter valores do Excel, removendo .0 de números inteiros
                    def converter_valor(valor):
                        if pd.notna(valor):
                            valor_str = str(valor).strip()
                            # Se termina com .0, remover (número inteiro lido como float pelo pandas)
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

    def corrigir_texto_hipotese_diagnostica(self, texto_original):
        """Aplica regras de correção para textos de hipótese diagnóstica"""
        if not texto_original or not texto_original.strip():
            return texto_original
        
        texto = texto_original.strip()
        texto_original_para_log = texto
        
        # Regras de correção para biópsias
        regras_biopsia = [
            # Biópsia Gástrica (com e sem acento, com e sem "de")
            (r'B[ií][oó]psia\s+G[aá]strica?', 'Biópsias Gástricas'),
            (r'B[ií][oó]psia\s+de\s+G[aá]strica?', 'Biópsias Gástricas'),
            (r'B[ií][oó]psia\s+g[aá]strica?', 'Biópsias Gástricas'),
            (r'B[ií][oó]psia\s+de\s+g[aá]strica?', 'Biópsias Gástricas'),
            # Biópsia de Próstata (com e sem acento, com e sem "de")
            (r'B[ií][oó]psia\s+de\s+Pr[oó]stata', 'Biópsias de Próstata'),
            (r'B[ií][oó]psia\s+Pr[oó]stata', 'Biópsias de Próstata'),
            (r'B[ií][oó]psia\s+de\s+pr[oó]stata', 'Biópsias de Próstata'),
            (r'B[ií][oó]psia\s+pr[oó]stata', 'Biópsias de Próstata'),
            # Biópsia de Reto (com e sem "de")
            (r'B[ií][oó]psia\s+de\s+Reto', 'Biópsias de Reto'),
            (r'B[ií][oó]psia\s+Reto', 'Biópsias de Reto'),
            (r'B[ií][oó]psia\s+de\s+reto', 'Biópsias de Reto'),
            (r'B[ií][oó]psia\s+reto', 'Biópsias de Reto'),
        ]
        
        # Regras de correção para lesões
        regras_lesao = [
            (r'Les[ãa]o\s+de\s+Reto', 'Lesões do Reto'),
            (r'Les[ãa]o\s+do\s+Reto', 'Lesões do Reto'),
            (r'Les[ãa]o\s+Reto', 'Lesões do Reto'),
            (r'Les[ãa]o\s+de\s+G[aá]strica', 'Lesões Gástricas'),
            (r'Les[ãa]o\s+do\s+G[aá]strica', 'Lesões Gástricas'),
            (r'Les[ãa]o\s+G[aá]strica', 'Lesões Gástricas'),
            (r'Les[ãa]o\s+de\s+Pr[oó]stata', 'Lesões de Próstata'),
            (r'Les[ãa]o\s+do\s+Pr[oó]stata', 'Lesões de Próstata'),
            (r'Les[ãa]o\s+Pr[oó]stata', 'Lesões de Próstata'),
        ]
        
        # Aplicar regras de biópsia
        for padrao, substituicao in regras_biopsia:
            if re.search(padrao, texto, re.IGNORECASE):
                texto = re.sub(padrao, substituicao, texto, flags=re.IGNORECASE)
                log_message(f"🔧 Correção aplicada (Biópsia): '{texto_original_para_log}' → '{texto}'", "INFO")
                break
        
        # Aplicar regras de lesão (só se não foi alterado por biópsia)
        if texto == texto_original_para_log:
            for padrao, substituicao in regras_lesao:
                if re.search(padrao, texto, re.IGNORECASE):
                    texto = re.sub(padrao, substituicao, texto, flags=re.IGNORECASE)
                    log_message(f"🔧 Correção aplicada (Lesão): '{texto_original_para_log}' → '{texto}'", "INFO")
                    break
        
        # Se houve alteração, log de sucesso
        if texto != texto_original_para_log:
            log_message(f"✅ Texto corrigido: '{texto_original_para_log}' → '{texto}'", "SUCCESS")
        else:
            log_message(f"ℹ️ Nenhuma correção necessária para: '{texto}'", "INFO")
        
        return texto

    def fazer_login_unimed(self, driver, wait, username, password):
        """Faz login no portal da Unimed"""
        log_message("Fazendo login no portal Unimed...", "INFO")
        driver.get("https://webmed.unimedlondrina.com.br/prestador/")
        
        # Aguardar e preencher campo usuário
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "operador")))
        campo_usuario.clear()
        campo_usuario.send_keys(username)
        
        # Preencher campo senha
        campo_senha = driver.find_element(By.ID, "senha")
        campo_senha.clear()
        campo_senha.send_keys(password)
        
        # Clicar em entrar
        botao_entrar = driver.find_element(By.ID, "entrar")
        botao_entrar.click()
        time.sleep(2.5)
        
        log_message("✅ Login realizado com sucesso", "SUCCESS")

    def acessar_pagina_procedimento(self, driver):
        """Acessa a página de procedimento específica da Unimed"""
        url_procedimento = "https://webmed.unimedlondrina.com.br/prestador/procedimento.php?pagina=ff25c04430244fa10de866898f1a24d2"
        log_message(f"Acessando página de procedimentos: {url_procedimento}", "INFO")
        driver.get(url_procedimento)
        time.sleep(3)
        log_message("✅ Página de procedimentos acessada", "SUCCESS")

    def verificar_erro_carteirinha(self, driver, wait):
        """Verifica se houve erro ao preencher o número da carteirinha"""
        try:
            # Guardar janela original
            janela_original = driver.current_window_handle
            
            # Aguardar um pouco para ver se popup abre
            time.sleep(2)
            
            # Verificar se há novas janelas
            todas_janelas = driver.window_handles
            
            if len(todas_janelas) > 1:
                # Há um popup aberto, verificar se é página de erro
                for janela in todas_janelas:
                    if janela != janela_original:
                        driver.switch_to.window(janela)
                        break
                
                # Verificar se é a página de erro
                url_atual = driver.current_url
                log_message(f"🔍 Popup detectado - URL: {url_atual}", "INFO")
                
                if "localizaUsuarioUnimed.php" in url_atual:
                    log_message("⚠️ Página de erro de carteirinha detectada", "WARNING")
                    
                    # Verificar mensagens de erro
                    mensagem_erro = None
                    
                    # 1. Verificar mensagem na div de erro
                    try:
                        elemento_erro = driver.find_element(By.CSS_SELECTOR, "#erro td")
                        if elemento_erro.is_displayed():
                            mensagem_erro = "Número de carteirinha inválido!"
                            log_message(f"❌ Erro detectado: {mensagem_erro}", "ERROR")
                    except Exception:
                        pass
                    
                    # 2. Verificar campo nome com mensagem de erro
                    try:
                        campo_nome = driver.find_element(By.ID, "nome")
                        valor_nome = campo_nome.get_attribute("value")
                        if valor_nome and "não encontrado" in valor_nome.lower():
                            mensagem_erro = valor_nome.strip()
                            log_message(f"❌ Erro no campo nome: {mensagem_erro}", "ERROR")
                    except Exception:
                        pass
                    
                    # Fechar popup
                    driver.close()
                    driver.switch_to.window(janela_original)
                    log_message("↩️ Popup de erro fechado, voltou para janela principal", "INFO")
                    
                    # Se encontrou erro, retornar
                    if mensagem_erro:
                        return {
                            'erro': True,
                            'mensagem': mensagem_erro
                        }
                else:
                    # Popup de outro tipo, voltar para janela original
                    driver.switch_to.window(janela_original)
            
            # Não há erro
            return {'erro': False}
            
        except Exception as e:
            log_message(f"⚠️ Erro ao verificar popup de erro: {e}", "WARNING")
            # Tentar voltar para janela original
            try:
                driver.switch_to.window(janela_original)
            except:
                pass
            return {'erro': False}
    
    def formatar_cartao_17_digitos(self, cartao):
        """Formata o número do cartão para ter 17 dígitos, adicionando zeros antes se necessário"""
        cartao_limpo = str(cartao).strip()
        
        # Remover ".0" se o Excel converteu para float
        if cartao_limpo.endswith('.0'):
            cartao_limpo = cartao_limpo[:-2]
            log_message(f"🔧 Removido '.0' do final do cartão: {cartao_limpo}", "INFO")
        
        # Remover apóstrofe do Excel e outros caracteres especiais, manter apenas números e letras
        cartao_sem_apostrofe = cartao_limpo.lstrip("'")  # Remove apóstrofe do início
        cartao_sem_espacos = ''.join(cartao_sem_apostrofe.split())
        
        # Remover qualquer ponto decimal remanescente (apenas números devem permanecer)
        cartao_apenas_numeros = re.sub(r'[^0-9]', '', cartao_sem_espacos)
        
        if len(cartao_apenas_numeros) < 17:
            # Adicionar zeros à esquerda para completar 17 dígitos
            zeros_necessarios = 17 - len(cartao_apenas_numeros)
            cartao_formatado = "0" * zeros_necessarios + cartao_apenas_numeros
            log_message(f"📋 Cartão formatado: '{cartao_limpo}' → {cartao_formatado} (17 dígitos)", "INFO")
            return cartao_formatado
        elif len(cartao_apenas_numeros) == 17:
            log_message(f"📋 Cartão já tem 17 dígitos: {cartao_apenas_numeros}", "INFO")
            return cartao_apenas_numeros
        else:
            log_message(f"⚠️ Cartão com mais de 17 dígitos ({len(cartao_apenas_numeros)}): {cartao_apenas_numeros}", "WARNING")
            return cartao_apenas_numeros

    def extrair_apenas_numeros(self, crm):
        """Extrai apenas os números do CRM, removendo letras"""
        apenas_numeros = re.sub(r'[^0-9]', '', str(crm))
        log_message(f"📋 CRM formatado: {crm} → {apenas_numeros}", "INFO")
        return apenas_numeros

    def comparar_nomes_medicos(self, nome_procurado, nome_encontrado):
        """Compara a similaridade entre dois nomes de médicos"""
        # Normalizar nomes (lower case, remover espaços extras e acentos, etc.)
        def normalizar(nome):
            return re.sub(r'\W+', '', nome.lower().strip())

        nome_procurado_norm = normalizar(nome_procurado)
        nome_encontrado_norm = normalizar(nome_encontrado)

        # Usar SequenceMatcher para comparar
        matcher = difflib.SequenceMatcher(None, nome_procurado_norm, nome_encontrado_norm)
        return matcher.ratio()

    def buscar_medico_solicitante(self, driver, wait, crm, nome_medico):
        """Busca o médico solicitante no popup da Unimed"""
        try:
            # Guardar janela original
            janela_original = driver.current_window_handle
            
            # 1. Clicar no botão de busca do solicitante
            log_message("🔍 Clicando no botão de busca do solicitante...", "INFO")
            botao_busca = wait.until(EC.element_to_be_clickable((By.ID, "busca_solicitante")))
            botao_busca.click()
            
            # 2. Aguardar nova janela abrir e fazer switch
            time.sleep(3)
            
            # Verificar se há novas janelas
            todas_janelas = driver.window_handles
            if len(todas_janelas) > 1:
                # Mudar para a nova janela (popup)
                for janela in todas_janelas:
                    if janela != janela_original:
                        driver.switch_to.window(janela)
                        break
                log_message("✅ Mudou para janela do popup", "INFO")
            else:
                log_message("✅ Popup aberto na mesma janela", "INFO")
            
            # 3. Extrair apenas números do CRM
            crm_numeros = self.extrair_apenas_numeros(crm)
            
            # Armazenar campo nome e conselho para reuso
            campo_conselho = wait.until(EC.presence_of_element_located((By.ID, "conselho")))
            campo_nome = wait.until(EC.presence_of_element_located((By.ID, "nome")))
            botao_localizar = wait.until(EC.element_to_be_clickable((By.ID, "localizar")))
            
            medico_selecionado = False
            tentativas = [
                {"nome": nome_medico, "log": "CRM e Nome Completo"},
                {"nome": "", "log": "Apenas CRM"} # Fallback: buscar apenas por CRM
            ]

            for tentativa_idx, tentativa_config in enumerate(tentativas):
                log_message(f"🔍 Tentativa {tentativa_idx + 1}: Buscando médico por {tentativa_config['log']}...", "INFO")

                # Re-localizar elementos para evitar StaleElementReferenceException
                campo_conselho = wait.until(EC.presence_of_element_located((By.ID, "conselho")))
                campo_nome = wait.until(EC.presence_of_element_located((By.ID, "nome")))
                botao_localizar = wait.until(EC.element_to_be_clickable((By.ID, "localizar")))

                # Limpar e preencher campos
                campo_conselho.clear()
                campo_conselho.send_keys(crm_numeros)
                campo_nome.clear()
                if tentativa_config['nome']:
                    campo_nome.send_keys(tentativa_config['nome'].upper())
                    log_message(f"📝 Preenchendo campo nome com: {tentativa_config['nome']}", "INFO")
                else:
                    log_message("📝 Campo nome deixado vazio para busca apenas por CRM", "INFO")
                
                botao_localizar.click()
                time.sleep(3) # Aguardar tabela carregar

                try:
                    tabela = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.table-hover tbody")))
                    linhas = tabela.find_elements(By.TAG_NAME, "tr") # Use TAG_NAME para pegar todas as linhas

                    if len(linhas) <= 1: # Apenas cabeçalho ou nenhum resultado
                        log_message(f"⚠️ Nenhum resultado encontrado para CRM: {crm_numeros} e nome: '{tentativa_config['nome']}'", "WARNING")
                        if tentativa_idx == len(tentativas) - 1: # Última tentativa e ainda sem resultados
                            raise Exception(f"Médico não encontrado após todas as tentativas para CRM: {crm}")
                        continue # Tentar próxima estratégia

                    log_message(f"✅ Encontrados {len(linhas) - 1} resultados para CRM: {crm_numeros} e nome: '{tentativa_config['nome']}'", "INFO")

                    # Processar resultados
                    melhor_similaridade = -1
                    melhor_medico = None

                    for linha in linhas[1:]: # Pula o cabeçalho
                        try:
                            colunas = linha.find_elements(By.TAG_NAME, "td")
                            if len(colunas) >= 2:
                                nome_na_tabela = colunas[1].text.strip()
                                documento_na_tabela = colunas[2].text.strip() if len(colunas) >= 3 else ""

                                log_message(f"📋 Médico na tabela: {nome_na_tabela} - {documento_na_tabela}", "INFO")
                                
                                # Se a busca foi com nome, verificar similaridade
                                if tentativa_config['nome']:
                                    similaridade = self.comparar_nomes_medicos(nome_medico, nome_na_tabela)
                                    if similaridade > melhor_similaridade:
                                        melhor_similaridade = similaridade
                                        melhor_medico = linha
                                    log_message(f"⚖️ Similaridade com '{nome_medico}': {similaridade:.2f}", "INFO")
                                else: # Se a busca foi apenas por CRM, o primeiro é o melhor (se houver)
                                    melhor_medico = linha
                                    break

                        except Exception as e: # Erro ao processar linha, continuar para a próxima
                            log_message(f"⚠️ Erro ao processar linha da tabela: {e}", "WARNING")
                            continue

                    if melhor_medico:
                        # Capturar nome antes do clique (evita stale caso a janela feche)
                        try:
                            tds_sel = melhor_medico.find_elements(By.TAG_NAME, "td")
                            nome_selecionado = tds_sel[1].text.strip() if len(tds_sel) > 1 else ""
                        except Exception:
                            nome_selecionado = ""
                        
                        # Clicar na linha do médico encontrado
                        melhor_medico.click()
                        log_message(f"✅ Médico selecionado: {nome_selecionado}", "SUCCESS")
                        medico_selecionado = True
                        
                        # Após o clique, o popup pode fechar. Garantir retorno para a janela original.
                        try:
                            time.sleep(1)
                            if janela_original in driver.window_handles:
                                driver.switch_to.window(janela_original)
                                log_message("↩️ Voltou para janela principal após selecionar médico", "INFO")
                        except Exception as e:
                            log_message(f"⚠️ Ao retornar para janela principal: {e}", "WARNING")
                        
                        break # Sai do loop de tentativas, pois o médico foi encontrado e selecionado
                    else:
                        log_message(f"⚠️ Nenhuma correspondência adequada encontrada para CRM: {crm_numeros} e nome: {nome_medico}", "WARNING")
                        if tentativa_idx == len(tentativas) - 1: # Última tentativa e ainda sem resultados
                            raise Exception(f"Médico não encontrado após todas as tentativas para CRM: {crm}")

                except Exception as e: # Erro no processamento da tabela
                    log_message(f"❌ Erro ao processar tabela de médicos: {e}", "ERROR")
                    if tentativa_idx == len(tentativas) - 1: # Última tentativa e ainda com erro
                        # Tentar voltar para janela original em caso de erro
                        try:
                            driver.switch_to.window(janela_original)
                            log_message("🔄 Voltou para janela principal após erro", "INFO")
                        except: # Caso a janela original não esteja mais acessível
                            pass
                        raise Exception(f"Falha na busca do médico: {e}")
                    continue # Tentar próxima estratégia

            if not medico_selecionado:
                raise Exception(f"Médico não foi selecionado após todas as tentativas para CRM: {crm}")

            # 9. Aguardar popup fechar automaticamente e voltar para janela original
            time.sleep(3)
            
            # O popup fecha automaticamente, então só precisamos voltar para janela original
            driver.switch_to.window(janela_original)
            log_message("✅ Médico selecionado, voltou para janela principal", "SUCCESS")

        except Exception as e:
            log_message(f"❌ Erro na busca do médico solicitante: {e}", "ERROR")
            # Garantir que volta para janela original em caso de erro
            try:
                if 'janela_original' in locals():
                    driver.switch_to.window(janela_original)
            except:
                pass
            raise e

    def preencher_campos_fixos(self, driver):
        """Preenche os campos fixos do formulário"""
        try:
            # Regime de atendimento: 01 - Ambulatorial
            log_message("📝 Preenchendo regime de atendimento: 01 - Ambulatorial", "INFO")
            js_regime = '''
            $("#regime_atendimento")
              .val("01")
              .trigger("change");
            '''
            driver.execute_script(js_regime)
            
            # Aguardar um pouco
            time.sleep(1)
            
            # Tipo de atendimento: 23 - Exame
            log_message("📝 Preenchendo tipo de atendimento: 23 - Exame", "INFO")
            js_tipo = '''
            $("#tipo_atendimento")
              .val("23")
              .trigger("change")
              .trigger("blur");
            '''
            driver.execute_script(js_tipo)
            
            time.sleep(1)
            log_message("✅ Campos fixos preenchidos", "SUCCESS")
            
        except Exception as e:
            log_message(f"❌ Erro ao preencher campos fixos: {e}", "ERROR")
            raise e

    def preencher_hipotese_diagnostica(self, driver, wait, texto):
        """Preenche a hipótese diagnóstica usando o campo TEXTO do Excel"""
        try:
            if not texto or not texto.strip():
                log_message("⚠️ Texto vazio, pulando hipótese diagnóstica", "WARNING")
                return
            
            texto_formatado = self.corrigir_texto_hipotese_diagnostica(texto)
            log_message(f"📝 Preenchendo hipótese diagnóstica: {texto_formatado}", "INFO")
            
            # 1. Clicar no campo select2 para abrir
            log_message("🔍 Clicando no campo de hipótese diagnóstica...", "INFO")
            select2_container = wait.until(EC.element_to_be_clickable((
                By.CSS_SELECTOR, 
                "#selecionarHipotese .select2-container .select2-selection"
            )))
            select2_container.click()
            
            # 2. Aguardar campo de busca aparecer
            time.sleep(2)
            
            # 3. Preencher campo de busca
            log_message(f"📝 Digitando texto: {texto_formatado}", "INFO")
            campo_busca = wait.until(EC.presence_of_element_located((
                By.CSS_SELECTOR, 
                ".select2-search__field"
            )))
            campo_busca.clear()
            campo_busca.send_keys(texto_formatado)
            
            # 4. Aguardar resultados carregar
            time.sleep(3)
            
            # 5. Verificar se há resultados ou se precisa usar "DIGITAR MANUALMENTE"
            try:
                # Procurar por opções de resultado
                resultados = driver.find_elements(By.CSS_SELECTOR, 
                    ".select2-results__options .select2-results__option")
                
                encontrou_resultado = False
                for resultado in resultados:
                    texto_resultado = resultado.text.strip()
                    
                    # Se encontrou "DIGITAR MANUALMENTE", clica nele
                    if "DIGITAR MANUALMENTE" in texto_resultado.upper():
                        log_message("📝 Selecionando 'DIGITAR MANUALMENTE'", "INFO")
                        resultado.click()
                        encontrou_resultado = True
                        break
                    # Se encontrou um resultado válido (não é mensagem de erro), pode clicar
                    elif (texto_resultado and 
                          "Digite 3 ou mais caracteres" not in texto_resultado and
                          "para selecionar" not in texto_resultado):
                        log_message(f"✅ Encontrou resultado: {texto_resultado}", "SUCCESS")
                        resultado.click()
                        encontrou_resultado = True
                        break
                
                if not encontrou_resultado:
                    log_message("⚠️ Nenhum resultado encontrado, tentando 'DIGITAR MANUALMENTE'", "WARNING")
                    # Tentar encontrar especificamente "DIGITAR MANUALMENTE"
                    digitar_manual = driver.find_element(By.XPATH, 
                        "//li[contains(text(), 'DIGITAR MANUALMENTE')]")
                    digitar_manual.click()
                
                time.sleep(1)
                log_message("✅ Hipótese diagnóstica preenchida", "SUCCESS")
                
            except Exception as e:
                log_message(f"⚠️ Erro ao selecionar hipótese: {e}", "WARNING")
                # Tentar fechar o dropdown em caso de erro
                try:
                    driver.execute_script("$('.select2-container').select2('close');")
                except:
                    pass
                
        except Exception as e:
            log_message(f"❌ Erro ao preencher hipótese diagnóstica: {e}", "ERROR")
            raise e

    def preencher_procedimentos(self, driver, procedimentos_str, quantidades_str):
        """Preenche os procedimentos e quantidades baseado nos dados do Excel"""
        try:
            if not procedimentos_str or not quantidades_str:
                log_message("⚠️ Procedimentos ou quantidades vazios", "WARNING")
                return
            
            # Processar strings dos procedimentos e quantidades
            procedimentos = [p.strip() for p in str(procedimentos_str).split(',')]
            quantidades = [q.strip() for q in str(quantidades_str).split(',')]
            
            # Verificar se as listas têm o mesmo tamanho
            if len(procedimentos) != len(quantidades):
                log_message(f"⚠️ Número de procedimentos ({len(procedimentos)}) difere do número de quantidades ({len(quantidades)})", "WARNING")
                # Ajustar para o menor tamanho
                min_size = min(len(procedimentos), len(quantidades))
                procedimentos = procedimentos[:min_size]
                quantidades = quantidades[:min_size]
            
            log_message(f"📋 Processando {len(procedimentos)} procedimentos:", "INFO")
            for i, (proc, qtd) in enumerate(zip(procedimentos, quantidades)):
                log_message(f"   {i}: {proc} = {qtd}", "INFO")
            
            # Preencher cada procedimento e quantidade
            for i, (procedimento, quantidade) in enumerate(zip(procedimentos, quantidades)):
                if i >= 5:  # Máximo de 5 campos (0 a 4)
                    log_message(f"⚠️ Limite de 5 procedimentos atingido, ignorando restantes", "WARNING")
                    break
                
                try:
                    # Preencher procedimento
                    log_message(f"📝 Preenchendo procedimento{i}: {procedimento}", "INFO")
                    js_procedimento = f'''
                    $("#procedimento{i}")
                      .val("{procedimento}")
                      .trigger("input")
                      .trigger("change") 
                      .trigger("blur");
                    '''
                    driver.execute_script(js_procedimento)
                    
                    # Aguardar um pouco
                    time.sleep(1)
                    
                    # Preencher quantidade
                    log_message(f"📝 Preenchendo quantidade{i}: {quantidade}", "INFO")
                    js_quantidade = f'''
                    $("#quantidade{i}")
                      .removeAttr("readonly")  
                      .val("{quantidade}")                
                      .trigger("input")
                      .trigger("change")
                      .trigger("blur");
                    '''
                    driver.execute_script(js_quantidade)
                    
                    time.sleep(1)
                    log_message(f"✅ Procedimento {i} preenchido: {procedimento} (qtd: {quantidade})", "SUCCESS")
                    
                except Exception as e:
                    log_message(f"❌ Erro ao preencher procedimento {i}: {e}", "ERROR")
                    continue
            
            log_message("✅ Procedimentos e quantidades preenchidos", "SUCCESS")
            
        except Exception as e:
            log_message(f"❌ Erro ao processar procedimentos: {e}", "ERROR")
            raise e

    def fechar_popup_aviso(self, driver, wait):
        """Fecha popup de aviso se aparecer (ex: aviso de cobertura contratual)"""
        try:
            # Verificar se há popup de aviso visível
            popups_aviso = driver.find_elements(By.CSS_SELECTOR, ".ui-dialog.ui-widget.ui-widget-content")
            
            for popup in popups_aviso:
                if popup.is_displayed():
                    # Verificar se é o aviso de cobertura ou outros avisos
                    try:
                        texto_popup = popup.text
                        
                        # Lista de mensagens de aviso que devem ser fechadas (não são erros)
                        avisos_conhecidos = [
                            "Sem cobertura contratual para Materiais e Medicamentos FORA de Ambiente Hospitalar",
                            "Atenção"
                        ]
                        
                        # Verificar se o popup contém alguma mensagem de aviso conhecida
                        eh_aviso = any(aviso.lower() in texto_popup.lower() for aviso in avisos_conhecidos)
                        
                        if eh_aviso:
                            log_message(f"ℹ️ Popup de aviso detectado: {texto_popup[:100]}...", "INFO")
                            
                            # Tentar clicar no botão Ok usando múltiplos seletores
                            botao_ok = None
                            seletores_ok = [
                                ".ui-dialog-buttonset button",
                                ".ui-dialog-buttonpane button",
                                "button[type='button']"
                            ]
                            
                            for seletor in seletores_ok:
                                try:
                                    botoes = popup.find_elements(By.CSS_SELECTOR, seletor)
                                    for botao in botoes:
                                        if botao.is_displayed() and 'ok' in botao.text.lower():
                                            botao_ok = botao
                                            break
                                    if botao_ok:
                                        break
                                except Exception:
                                    continue
                            
                            if botao_ok:
                                botao_ok.click()
                                time.sleep(0.5)
                                log_message("✅ Popup de aviso fechado, continuando processamento", "SUCCESS")
                            else:
                                # Se não encontrou botão, tentar fechar via JavaScript
                                log_message("⚠️ Botão Ok não encontrado, tentando fechar via JavaScript...", "WARNING")
                                driver.execute_script("arguments[0].remove();", popup)
                                time.sleep(0.5)
                    except Exception as e:
                        log_message(f"⚠️ Erro ao processar popup de aviso: {e}", "WARNING")
                        continue
                        
        except Exception as e:
            # Se falhar, não é crítico - apenas logar
            log_message(f"⚠️ Erro ao verificar popups de aviso: {e}", "WARNING")
    
    def autorizar_guia(self, driver, wait):
        """Clica no botão Autorizar e captura o resultado"""
        try:
            log_message("🔄 Clicando no botão Autorizar...", "INFO")
            
            # Tentar fechar overlays que possam estar bloqueando o botão
            try:
                driver.execute_script("""
                    // Fechar overlays jQuery UI
                    $('.ui-widget-overlay').remove();
                    $('.ui-front').remove();
                """)
                time.sleep(0.5)
            except Exception:
                pass
            
            # Clicar no botão autorizar - tentar múltiplas vezes se necessário
            botao_autorizar = wait.until(EC.element_to_be_clickable((By.ID, "autorizar")))
            
            # Tentar clicar normalmente primeiro
            try:
                botao_autorizar.click()
            except Exception as e:
                # Se falhar, tentar com JavaScript
                log_message("⚠️ Clique normal falhou, tentando com JavaScript...", "WARNING")
                driver.execute_script("arguments[0].click();", botao_autorizar)
            
            # Aguardar um momento e verificar se há popups de aviso
            time.sleep(2)
            self.fechar_popup_aviso(driver, wait)
            
            # Aguardar modal de resultado aparecer
            time.sleep(1)
            
            # Verificar se modal de sucesso apareceu
            try:
                # Aguardar modal aparecer
                modal = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#form-1.modal")))
                
                # Verificar se é modal de sucesso ou erro
                modal_body = modal.find_element(By.CSS_SELECTOR, ".modal-body")
                texto_modal = modal_body.text.strip()
                
                log_message(f"📋 Texto do modal: {texto_modal}", "INFO")
                
                # Procurar pelo número da guia no texto
                import re
                
                # Padrão 1: Guia gerada com sucesso
                match_guia_sucesso = re.search(r'Guia\s+(\d+)\s+gerada com sucesso', texto_modal)
                
                # Padrão 2: Autorização enviada para análise (também gera número)
                match_guia_analise = re.search(r'Autorização não concedida, porém enviada para análise\.\s*Número:\s*(\d+)', texto_modal)
                
                if match_guia_sucesso:
                    numero_guia = match_guia_sucesso.group(1)
                    log_message(f"✅ Guia gerada com sucesso: {numero_guia}", "SUCCESS")
                    
                    # Clicar no botão OK para fechar modal
                    try:
                        botao_ok = modal.find_element(By.ID, "btn_OK1")
                        botao_ok.click()
                        time.sleep(1)
                        log_message("✅ Modal fechado", "INFO")
                    except Exception as e:
                        log_message(f"⚠️ Erro ao fechar modal: {e}", "WARNING")
                    
                    return {
                        'sucesso': True,
                        'numero_guia': numero_guia,
                        'mensagem': f"Guia {numero_guia} gerada com sucesso"
                    }
                elif match_guia_analise:
                    numero_guia = match_guia_analise.group(1)
                    log_message(f"⚠️ Guia enviada para análise: {numero_guia}", "WARNING")
                    log_message("ℹ️ Tratando como sucesso parcial - guia será validada posteriormente", "INFO")
                    
                    # Clicar no botão OK para fechar modal
                    try:
                        botao_ok = modal.find_element(By.ID, "btn_OK1")
                        botao_ok.click()
                        time.sleep(1)
                        log_message("✅ Modal fechado", "INFO")
                    except Exception as e:
                        log_message(f"⚠️ Erro ao fechar modal: {e}", "WARNING")
                    
                    return {
                        'sucesso': True,
                        'numero_guia': numero_guia,
                        'mensagem': f"Guia {numero_guia} enviada para análise - aguardando aprovação",
                        'status_especial': 'analise'
                    }
                else:
                    # Se não encontrou nenhum padrão conhecido, é erro real
                    log_message(f"❌ Erro no lançamento da guia: {texto_modal}", "ERROR")
                    
                    # Tentar fechar modal
                    try:
                        botao_ok = modal.find_element(By.ID, "btn_OK1")
                        botao_ok.click()
                        time.sleep(1)
                    except:
                        pass
                    
                    return {
                        'sucesso': False,
                        'numero_guia': None,
                        'mensagem': texto_modal
                    }
                    
            except Exception as e:
                log_message(f"❌ Erro ao processar modal: {e}", "ERROR")
                return {
                    'sucesso': False,
                    'numero_guia': None,
                    'mensagem': f"Erro ao processar modal: {e}"
                }
                
        except Exception as e:
            log_message(f"❌ Erro ao autorizar guia: {e}", "ERROR")
            return {
                'sucesso': False,
                'numero_guia': None,
                'mensagem': f"Erro ao autorizar: {e}"
            }

    def consultar_status_guia(self, driver, wait, numero_guia):
        """Consulta o status de uma guia na página de rastreabilidade"""
        try:
            log_message(f"🔍 Consultando status da guia {numero_guia}...", "INFO")
            
            # Acessar página de rastreabilidade
            url_rastreabilidade = "https://webmed.unimedlondrina.com.br/prestador/Rastreabilidade.php"
            driver.get(url_rastreabilidade)
            time.sleep(3)
            
            # Preencher campo do número da guia
            log_message(f"📝 Preenchendo número da guia: {numero_guia}", "INFO")
            campo_numero_guia = wait.until(EC.presence_of_element_located((By.ID, "numeroGuia")))
            campo_numero_guia.clear()
            campo_numero_guia.send_keys(str(numero_guia))
            
            # Clicar no botão consultar
            log_message("🔍 Clicando no botão Consultar...", "INFO")
            botao_consultar = wait.until(EC.element_to_be_clickable((By.ID, "consultar")))
            botao_consultar.click()
            
            # Aguardar carregamento da página
            time.sleep(3)
            
            # Tentar extrair o status da guia
            try:
                campo_status = wait.until(EC.presence_of_element_located((By.ID, "status")))
                status_guia = campo_status.get_attribute("value").strip()
                
                log_message(f"✅ Status da guia {numero_guia}: {status_guia}", "SUCCESS")
                
                return {
                    'sucesso': True,
                    'status_guia': status_guia,
                    'numero_guia': numero_guia
                }
                
            except Exception as e:
                log_message(f"⚠️ Não foi possível obter o status da guia {numero_guia}: {e}", "WARNING")
                return {
                    'sucesso': False,
                    'status_guia': 'Erro ao consultar',
                    'numero_guia': numero_guia,
                    'erro': str(e)
                }
                
        except Exception as e:
            log_message(f"❌ Erro ao consultar status da guia {numero_guia}: {e}", "ERROR")
            return {
                'sucesso': False,
                'status_guia': 'Erro ao consultar',
                'numero_guia': numero_guia,
                'erro': str(e)
            }

    def fazer_login_pathoweb(self, driver, wait, username, password):
        """Faz login no PathoWeb e navega para o módulo de faturamento"""
        try:
            log_message("🔐 Fazendo login no PathoWeb...", "INFO")
            
            # URL do PathoWeb
            url = "https://dap.pathoweb.com.br/login/auth"
            driver.get(url)
            
            # Preencher credenciais
            wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
            driver.find_element(By.ID, "password").send_keys(password)
            botao_submit = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            self.click_element(driver, botao_submit, "botão login PathoWeb")
            
            log_message("Verificando se precisa navegar para módulo de faturamento...", "INFO")
            current_url = driver.current_url

            if current_url == "https://dap.pathoweb.com.br/" or "trocarModulo" in current_url:
                log_message("Detectada tela de seleção de módulos - navegando para módulo de faturamento...", "INFO")
                try:
                    modulo_link = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                        "a[href='/site/trocarModulo?modulo=2']", condition="presence")
                    self.click_element(driver, modulo_link, "link módulo faturamento")
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

            # Fechar modal se aparecer
            try:
                modal_close_button = driver.find_element(By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button")
                # Em modo headless, não verificar is_displayed()
                if self.headless_mode or modal_close_button.is_displayed():
                    self.click_element(driver, modal_close_button, "fechar modal inicial")
                    time.sleep(1)
            except Exception:
                pass

            # Acessar explicitamente a página do módulo de faturamento
            log_message("Acessando módulo de faturamento via URL...", "INFO")
            driver.get("https://dap.pathoweb.com.br/moduloFaturamento/index")

            # Clicar no botão "Preparar exames para fatura"
            log_message("Clicando em 'Preparar exames para fatura'...", "INFO")
            try:
                preparar_btn = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                    "a.btn.btn-danger.chamadaAjax.setupAjax[data-url='/moduloFaturamento/preFaturamento']",
                    condition="presence")
                self.click_element(driver, preparar_btn, "botão 'Preparar exames para fatura'")
            except Exception:
                preparar_btn = self.wait_for_element(driver, wait, By.XPATH,
                    "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]",
                    condition="presence")
                self.click_element(driver, preparar_btn, "botão 'Preparar exames para fatura' (alternativo)")

            # Aguardar possível spinner/modal carregar
            try:
                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                log_message("✅ Modal de carregamento fechado", "INFO")
            except Exception:
                time.sleep(1)

            log_message("✅ Login no PathoWeb realizado e página de pré-faturamento acessada", "SUCCESS")
            return True
            
        except Exception as e:
            log_message(f"❌ Erro ao fazer login no PathoWeb: {e}", "ERROR")
            return False

    def preencher_campos_exame(self, driver, wait, numero_guia_unimed):
        """Preenche os campos de data e número da guia no modal do exame"""
        try:
            log_message("📝 Preenchendo campos do exame...", "INFO")
            
            # Data atual no formato necessário
            from datetime import datetime
            data_atual = datetime.now()
            ymd = data_atual.strftime("%Y-%m-%d")
            br = data_atual.strftime("%d/%m/%Y")
            
            log_message(f"📅 Data a ser preenchida: {br}", "INFO")
            log_message(f"🔢 Número da guia a ser preenchido: {numero_guia_unimed}", "INFO")
            
            # Aguardar um pouco para garantir que a tabela está carregada
            time.sleep(2)
            
            # Preencher campo de data de autorização usando jQuery
            log_message("📝 Preenchendo data de autorização...", "INFO")
            js_data_autorizacao = f'''
            const $input = $('#requisicao_r input[name="dataAutorizacao"]').first();
            if ($input.length === 0) return;

            const $a = $input.closest('td').children('a.table-editable-ancora').first();

            $input
              .val('{ymd}')
              .attr('value', '{ymd}')
              .trigger('focus')
              .trigger('input')
              .trigger('change')
              .trigger('blur');

            if ($a.length) {{
              $a.text('{br}').css('display', 'inline');
            }}
            '''
            
            driver.execute_script(js_data_autorizacao)
            log_message("✅ Data de autorização preenchida", "SUCCESS")
            time.sleep(1)
            
            # Preencher campo de data de requisição usando jQuery
            log_message("📝 Preenchendo data de requisição...", "INFO")
            js_data_requisicao = f'''
            const $input = $('#requisicao_r input[name="dataRequisicao"]').first();
            if ($input.length === 0) return;

            const $a = $input.closest('td').children('a.table-editable-ancora').first();

            $input
              .val('{ymd}')
              .attr('value', '{ymd}')
              .trigger('focus')
              .trigger('input')
              .trigger('change')
              .trigger('blur');

            if ($a.length) {{
              $a.text('{br}').css('display', 'inline');
            }}
            '''
            
            driver.execute_script(js_data_requisicao)
            log_message("✅ Data de requisição preenchida", "SUCCESS")
            time.sleep(1)
            
            # Preencher número da guia usando a função jQuery que você forneceu
            log_message("📝 Preenchendo número da guia...", "INFO")
            js_numero_guia = f'''
            function typeNumeroGuia(texto, delay = 40) {{
              const $inp = $("#numeroGuiaInput");
              const $a   = $inp.closest('td').children('a.table-editable-ancora').first();

              // limpa antes
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
                  // consolida valor nos atributos e dispara change/blur (para AJAX no blur)
                  $inp.attr("value", texto)
                      .data("previous-value", texto)
                      .trigger("change")
                      .trigger("blur");
                }}
              }}, delay);
            }}

            // uso:
            typeNumeroGuia("{numero_guia_unimed}", 30);
            '''
            
            driver.execute_script(js_numero_guia)
            log_message(f"✅ Número da guia {numero_guia_unimed} preenchido", "SUCCESS")
            
            # Aguardar um pouco para o processamento
            time.sleep(3)
            
            # 1. Clicar no botão "Próximo" para salvar os dados do exame
            log_message("🔄 Clicando no botão 'Próximo' para salvar...", "INFO")
            try:
                botao_proximo = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                    "a.btn.btn-sm.btn-primary.wizardControl.chamadaAjax.setupAjax[data-url='/moduloFaturamento/saveAjaxExameParaFaturamento']",
                    condition="presence")
                self.click_element(driver, botao_proximo, "botão 'Próximo'")
                log_message("✅ Botão 'Próximo' clicado", "SUCCESS")
                
                # Aguardar processamento
                time.sleep(3)
                
            except Exception as e:
                log_message(f"⚠️ Erro ao clicar no botão 'Próximo': {e}", "WARNING")
                # Tentar encontrar o botão com seletor alternativo
                try:
                    botao_proximo_alt = self.wait_for_element(driver, wait, By.XPATH,
                        "//a[contains(@class, 'wizardControl') and contains(text(), 'Próximo')]",
                        condition="presence")
                    self.click_element(driver, botao_proximo_alt, "botão 'Próximo' (alternativo)")
                    log_message("✅ Botão 'Próximo' clicado (seletor alternativo)", "SUCCESS")
                    time.sleep(3)
                except Exception as e2:
                    log_message(f"❌ Erro ao clicar no botão 'Próximo' (tentativa alternativa): {e2}", "ERROR")
            
            # 2. Clicar no botão "Salvar" para finalizar
            log_message("💾 Clicando no botão 'Salvar' para finalizar...", "INFO")
            try:
                botao_salvar = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                    "a.btn.btn-sm.btn-primary.chamadaAjax.setupAjax[data-url='/moduloFaturamento/saveExameDadosClinicos']",
                    condition="presence")
                self.click_element(driver, botao_salvar, "botão 'Salvar'")
                log_message("✅ Botão 'Salvar' clicado", "SUCCESS")
                
                # Aguardar processamento
                time.sleep(3)
            except Exception as e:
                log_message(f"⚠️ Erro ao clicar no botão 'Salvar': {e}", "WARNING")
                # Tentar encontrar o botão com seletor alternativo
                try:
                    botao_salvar_alt = self.wait_for_element(driver, wait, By.XPATH,
                        "//a[contains(@class, 'chamadaAjax') and contains(text(), 'Salvar')]",
                        condition="presence")
                    self.click_element(driver, botao_salvar_alt, "botão 'Salvar' (alternativo)")
                    log_message("✅ Botão 'Salvar' clicado (seletor alternativo)", "SUCCESS")
                    time.sleep(3)
                except Exception as e2:
                    log_message(f"❌ Erro ao clicar no botão 'Salvar' (tentativa alternativa): {e2}", "ERROR")

            # Fechar o modal após salvar
            try:
                modal = wait.until(EC.presence_of_element_located((By.ID, "myModal")))
                try:
                    close_btn = modal.find_element(By.CSS_SELECTOR, "button.close[data-dismiss='modal']")
                except Exception:
                    close_btn = driver.find_element(By.CSS_SELECTOR, "#myModal button.close, #myModal .modal-header button.close")
                self.click_element(driver, close_btn, "botão fechar modal")
                time.sleep(1)
                log_message("✅ Modal fechado após salvar", "INFO")
            except Exception as e:
                log_message(f"⚠️ Não foi possível fechar o modal automaticamente: {e}", "WARNING")
            
            # Marcar todos os exames da tabela como 'Conferido'
            try:
                # Garantir que a tabela está presente
                wait.until(EC.presence_of_element_located((By.ID, "tabelaPreFaturamentoTbody")))
                self.marcar_exames_como_conferidos(driver, wait)
            except Exception as e:
                log_message(f"⚠️ Erro ao marcar exames como 'Conferido': {e}", "WARNING")

            log_message("✅ Campos do exame preenchidos e salvos com sucesso", "SUCCESS")
            return True
            
        except Exception as e:
            log_message(f"❌ Erro ao preencher campos do exame: {e}", "ERROR")
            return False

    def abrir_exame_pathoweb(self, driver, wait, numero_guia_original, numero_guia_unimed=None):
        """Abre um exame específico no PathoWeb usando o número da guia original"""
        try:
            log_message(f"🔍 Abrindo exame {numero_guia_original} no PathoWeb...", "INFO")
            
            # Digitar o código de barras no campo codigoBarras
            log_message(f"📝 Digitando código de barras: {numero_guia_original}", "INFO")
            campo_exame = wait.until(EC.element_to_be_clickable((By.ID, "codigoBarras")))

            # Limpar e preencher o campo
            campo_exame.clear()
            time.sleep(0.5)
            campo_exame.send_keys(str(numero_guia_original))
            log_message(f"✅ Código de barras {numero_guia_original} digitado no campo", "SUCCESS")
            time.sleep(0.5)
            
            # Clicar no botão Pesquisar
            pesquisar_btn = self.wait_for_element(driver, wait, By.ID, "pesquisaFaturamento", condition="presence")
            self.click_element(driver, pesquisar_btn, "botão Pesquisar")
            log_message("🔍 Pesquisando exame...", "INFO")
            
            # Aguardar carregamento dos resultados
            try:
                # Aguardar spinner se existir
                try:
                    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                    log_message("🔄 Carregando resultados...", "INFO")
                    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                except Exception:
                    time.sleep(5)
            except Exception:
                log_message("Tempo de carregamento excedido, verificando resultados mesmo assim...", "WARNING")
            
            # Aguardar mais um pouco para garantir que a tabela foi carregada
            time.sleep(3)
            
            # Verificar se há resultados
            tbody_rows = []
            selectors = [
                "#tabelaPreFaturamentoTbody tr",
                ".table-responsive table tbody tr",
                "table.table-striped tbody tr",
                "table.footable tbody tr"
            ]
            
            for selector in selectors:
                try:
                    tbody_rows = driver.find_elements(By.CSS_SELECTOR, selector)
                    if len(tbody_rows) > 0:
                        log_message(f"✅ Tabela de resultados encontrada usando seletor: {selector}", "SUCCESS")
                        break
                except Exception:
                    continue
            
            if len(tbody_rows) == 0:
                log_message(f"⚠️ Nenhum resultado encontrado para {numero_guia_original}", "WARNING")
                return False
            
            log_message(f"✅ Encontrados {len(tbody_rows)} resultados para a guia {numero_guia_original}", "SUCCESS")
            
            # Marcar checkbox do primeiro exame e clicar no botão "Abrir exame"
            log_message("📝 Marcando checkbox do primeiro exame...", "INFO")
            
            try:
                checkbox = tbody_rows[0].find_element(By.CSS_SELECTOR, "input[type='checkbox'][name='exameId']")
                if not checkbox.is_selected():
                    self.click_element(driver, checkbox, "checkbox do exame")
                    log_message("✅ Checkbox do exame marcado", "SUCCESS")
                else:
                    log_message("ℹ️ Checkbox já estava marcado", "INFO")
                
                time.sleep(1)
                
                # Procurar e clicar no botão "Abrir exame"
                log_message("🔍 Procurando botão 'Abrir exame'...", "INFO")
                
                abrir_btn = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                    "a.btn.btn-sm.btn-primary.chamadaAjax.toogleInicial.setupAjax[data-url='/moduloFaturamento/abrirExameCorrecao']",
                    condition="presence")
                log_message("✅ Botão 'Abrir exame' encontrado", "SUCCESS")
                
                # Clicar no botão
                self.click_element(driver, abrir_btn, "botão 'Abrir exame'")
                log_message("✅ Clique no botão 'Abrir exame' realizado", "SUCCESS")
                
                # Aguardar o modal aparecer
                log_message("⏳ Aguardando modal do exame abrir...", "INFO")
                time.sleep(3)
                
                # Verificar se o modal foi aberto
                try:
                    modal = wait.until(EC.presence_of_element_located((By.ID, "myModal")))
                    # Em modo headless, não verificar is_displayed() pois pode retornar False
                    if self.headless_mode or modal.is_displayed():
                        log_message("✅ Modal do exame aberto com sucesso", "SUCCESS")
                        
                        # Preencher campos do exame se número da guia Unimed foi fornecido
                        if numero_guia_unimed:
                            log_message("📝 Preenchendo campos do exame...", "INFO")
                            self.preencher_campos_exame(driver, wait, numero_guia_unimed)
                        else:
                            log_message("ℹ️ Número da guia Unimed não fornecido, pulando preenchimento", "INFO")
                        
                        return True
                    else:
                        log_message("⚠️ Modal encontrado mas não está visível", "WARNING")
                        time.sleep(2)
                        return True
                except Exception:
                    log_message("⚠️ Modal não encontrado, tentando continuar...", "WARNING")
                    time.sleep(2)
                    return True
                    
            except Exception as e:
                log_message(f"❌ Erro ao abrir exame: {e}", "ERROR")
                return False
                
        except Exception as e:
            log_message(f"❌ Erro ao abrir exame {numero_guia_original}: {e}", "ERROR")
            return False

    def limpar_mensagem_erro(self, mensagem):
        """Remove stacktraces técnicas e retorna mensagem limpa para o usuário"""
        if not mensagem:
            return ''
        
        # Converter para string
        msg_str = str(mensagem)
        
        # Se é um erro do Selenium, extrair apenas a mensagem principal
        if 'Stacktrace:' in msg_str or 'Session info:' in msg_str:
            # Pegar apenas a primeira linha (mensagem principal)
            linhas = msg_str.split('\n')
            for linha in linhas:
                # Pular linhas vazias e de context
                if linha.strip() and not linha.strip().startswith('('):
                    # Simplificar mensagem
                    if 'element click intercepted' in linha.lower():
                        return 'Erro ao clicar no botão (elemento sobreposto)'
                    elif 'not clickable' in linha.lower():
                        return 'Erro: elemento não clicável'
                    elif 'timeout' in linha.lower():
                        return 'Erro: tempo limite excedido'
                    elif 'no such element' in linha.lower():
                        return 'Erro: elemento não encontrado'
                    else:
                        # Retornar primeira linha limpa
                        return linha.strip()[:200]  # Limitar tamanho
        
        # Se não é erro do Selenium, retornar mensagem original (limitada)
        return msg_str[:200]  # Limitar a 200 caracteres
    
    def salvar_resultados_excel(self, excel_file, resultados):
        """Salva os resultados no arquivo Excel original"""
        try:
            log_message("💾 Salvando resultados no Excel...", "INFO")
            log_message(f"📁 Arquivo de destino: {excel_file}", "INFO")
            
            # Ler o arquivo Excel original com dtype específico para evitar conversão de números grandes
            df = pd.read_excel(excel_file, header=0, dtype={'CARTAO': str})
            log_message(f"📋 Colunas originais: {list(df.columns)}", "INFO")
            
            # Preservar nomes originais das colunas, mas criar um mapeamento para comparação
            colunas_originais = list(df.columns)
            mapa_colunas = {}
            for col in colunas_originais:
                col_normalizada = col.upper().strip()
                mapa_colunas[col_normalizada] = col
            
            log_message(f"📋 Mapeamento de colunas criado", "INFO")
            
            # Adicionar colunas de resultado se não existirem (preservando case original quando possível)
            colunas_adicionadas = []
            if 'NUMERO_GUIA' not in [c.upper().strip() for c in df.columns]:
                df['Numero_Guia'] = ''
                colunas_adicionadas.append('Numero_Guia')
            if 'STATUS_PROCESSAMENTO' not in [c.upper().strip() for c in df.columns]:
                df['Status_Processamento'] = ''
                colunas_adicionadas.append('Status_Processamento')
            if 'STATUS_GUIA' not in [c.upper().strip() for c in df.columns]:
                df['Status_Guia'] = ''
                colunas_adicionadas.append('Status_Guia')
            if 'MENSAGEM_ERRO' not in [c.upper().strip() for c in df.columns]:
                df['Mensagem_Erro'] = ''
                colunas_adicionadas.append('Mensagem_Erro')
            if 'DATA_PROCESSAMENTO' not in [c.upper().strip() for c in df.columns]:
                df['Data_Processamento'] = ''
                colunas_adicionadas.append('Data_Processamento')
            
            if colunas_adicionadas:
                log_message(f"✅ Colunas adicionadas: {colunas_adicionadas}", "SUCCESS")
            else:
                log_message("ℹ️ Todas as colunas já existiam", "INFO")
            
            log_message(f"📋 Colunas finais: {list(df.columns)}", "INFO")
            
            # Encontrar a coluna GUIA (pode ter qualquer case)
            coluna_guia = None
            for col in df.columns:
                if col.upper().strip() == 'GUIA':
                    coluna_guia = col
                    break
            
            if not coluna_guia:
                raise ValueError("Coluna GUIA não encontrada no Excel")
            
            # Definir nomes das colunas de resultado (usar existentes ou novos)
            def get_nome_coluna(nome_procurado):
                for col in df.columns:
                    if col.upper().strip() == nome_procurado.upper().strip():
                        return col
                # Se não encontrou, usar o nome padrão criado acima
                mapeamento = {
                    'NUMERO_GUIA': 'Numero_Guia',
                    'STATUS_PROCESSAMENTO': 'Status_Processamento', 
                    'STATUS_GUIA': 'Status_Guia',
                    'MENSAGEM_ERRO': 'Mensagem_Erro',
                    'DATA_PROCESSAMENTO': 'Data_Processamento'
                }
                return mapeamento.get(nome_procurado.upper().strip(), nome_procurado)
            
            col_numero_guia = get_nome_coluna('NUMERO_GUIA')
            col_status_proc = get_nome_coluna('STATUS_PROCESSAMENTO')
            col_status_guia = get_nome_coluna('STATUS_GUIA')
            col_mensagem_erro = get_nome_coluna('MENSAGEM_ERRO')
            col_data_proc = get_nome_coluna('DATA_PROCESSAMENTO')
            
            # Atualizar resultados
            for resultado in resultados:
                guia = resultado.get('guia')
                log_message(f"📝 Processando resultado para guia: {guia}", "INFO")
                
                # Encontrar linha correspondente
                mask = df[coluna_guia].astype(str).str.strip() == str(guia).strip()
                indices = df[mask].index
                
                if len(indices) > 0:
                    indice = indices[0]
                    log_message(f"✅ Linha encontrada no índice: {indice}", "SUCCESS")
                    
                    # Atualizar dados
                    df.loc[indice, col_data_proc] = resultado.get('timestamp', '')
                    df.loc[indice, col_status_guia] = resultado.get('status_guia', '')
                    
                    # Sempre preservar o número da guia se existir
                    df.loc[indice, col_numero_guia] = resultado.get('numero_guia', '')
                    
                    if resultado.get('status') == 'sucesso':
                        df.loc[indice, col_status_proc] = 'SUCESSO'
                        df.loc[indice, col_mensagem_erro] = ''
                    elif resultado.get('status') == 'analise':
                        # Guia criada mas enviada para análise/auditoria
                        df.loc[indice, col_status_proc] = 'ANÁLISE'
                        df.loc[indice, col_mensagem_erro] = resultado.get('mensagem', 'Guia enviada para análise')
                    else:
                        # Status de erro - preservar número da guia se houver
                        # (ex: guia criada mas não liberada após tentativas)
                        df.loc[indice, col_status_proc] = 'ERRO'
                        # Limpar mensagem de erro antes de salvar
                        mensagem_erro_original = resultado.get('erro', resultado.get('mensagem', ''))
                        mensagem_erro_limpa = self.limpar_mensagem_erro(mensagem_erro_original)
                        df.loc[indice, col_mensagem_erro] = mensagem_erro_limpa
                    
                    log_message(f"📝 Resultado salvo para guia {guia}: {df.loc[indice, col_status_proc]}", "INFO")
                    log_message(f"   - {col_numero_guia}: {df.loc[indice, col_numero_guia]}", "INFO")
                    log_message(f"   - {col_status_guia}: {df.loc[indice, col_status_guia]}", "INFO")
                else:
                    log_message(f"⚠️ Linha não encontrada para guia: {guia}", "WARNING")
            
            # Função auxiliar para salvar Excel com formatação de texto para CARTAO
            def salvar_excel_com_formato(caminho_arquivo, dataframe):
                """Salva Excel garantindo que CARTAO seja salvo como texto"""
                from openpyxl.utils import get_column_letter
                
                with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
                    dataframe.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                    # Obter a worksheet
                    worksheet = writer.sheets['Sheet1']
                    
                    # Encontrar coluna CARTAO
                    coluna_cartao_idx = None
                    for idx, col_name in enumerate(dataframe.columns, start=1):
                        if col_name.upper().strip() in ['CARTAO', 'CARTÃO']:
                            coluna_cartao_idx = idx
                            break
                    
                    # Se encontrou coluna CARTAO, formatar como texto
                    if coluna_cartao_idx:
                        col_letter = get_column_letter(coluna_cartao_idx)
                        
                        # Formatar toda a coluna como texto
                        for row in range(2, len(dataframe) + 2):  # +2 porque começa na linha 2 (depois do cabeçalho)
                            cell = worksheet[f'{col_letter}{row}']
                            cell.number_format = '@'  # @ = formato de texto
                            
                            # Se o valor atual tem notação científica, reconverter para texto
                            if cell.value:
                                valor_str = str(cell.value)
                                if 'E+' in valor_str or 'e+' in valor_str:
                                    # Tentar converter de notação científica para texto
                                    try:
                                        numero = float(valor_str)
                                        cell.value = f"{numero:.0f}"
                                    except:
                                        pass
                        
                        log_message(f"✅ Coluna CARTAO formatada como texto na coluna {col_letter}", "SUCCESS")
            
            # Estratégia 1: Tentar salvar no arquivo original (sobrescrever)
            try:
                log_message("💾 Tentando salvar no arquivo original...", "INFO")
                salvar_excel_com_formato(excel_file, df)
                log_message(f"✅ Resultados salvos no arquivo original: {excel_file}", "SUCCESS")
                return excel_file
                
            except PermissionError:
                log_message("⚠️ Arquivo original em uso, criando arquivo separado...", "WARNING")
            except Exception as e:
                log_message(f"⚠️ Erro ao salvar no arquivo original: {e}", "WARNING")
            
            # Estratégia 2: Criar arquivo separado com timestamp
            import os
            from datetime import datetime
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_base = os.path.splitext(excel_file)[0]
            nome_arquivo_separado = f"{nome_base}_com_resultados_{timestamp}.xlsx"
            
            try:
                salvar_excel_com_formato(nome_arquivo_separado, df)
                log_message(f"✅ Resultados salvos em arquivo separado: {nome_arquivo_separado}", "SUCCESS")
                return nome_arquivo_separado
                
            except Exception as e:
                log_message(f"❌ Erro ao salvar arquivo separado: {e}", "ERROR")
                
                # Estratégia 3: Fallback para CSV
                try:
                    nome_csv = f"{nome_base}_resultados_{timestamp}.csv"
                    df.to_csv(nome_csv, index=False, encoding='utf-8-sig')
                    log_message(f"✅ Resultados salvos como CSV: {nome_csv}", "SUCCESS")
                    return nome_csv
                except Exception as e3:
                    log_message(f"❌ Erro ao salvar CSV: {e3}", "ERROR")
                    return None
            
        except Exception as e:
            log_message(f"❌ Erro geral ao salvar resultados: {e}", "ERROR")
            return None

    def processar_guia_unimed(self, driver, wait, dados):
        """Processa uma guia individual na página da Unimed"""
        try:
            log_message(f"🔄 Iniciando processamento da guia {dados['guia']}", "INFO")
            
            # Logar os dados que serão processados
            log_message(f"📝 Dados a processar:", "INFO")
            log_message(f"   - Guia: {dados['guia']}", "INFO")
            log_message(f"   - Cartão: {dados['cartao']}", "INFO")
            log_message(f"   - Médico: {dados['medico']}", "INFO")
            log_message(f"   - CRM: {dados['crm']}", "INFO")
            log_message(f"   - Procedimentos: {dados['procedimentos']}", "INFO")
            log_message(f"   - Quantidade: {dados['qtd']}", "INFO")
            log_message(f"   - Texto: {dados['texto'][:50]}..." if dados['texto'] else "   - Texto: (vazio)", "INFO")
            
            # Verificar se o cartão está vazio
            if not dados['cartao'] or not str(dados['cartao']).strip():
                log_message(f"⚠️ Cartão vazio para a guia {dados['guia']}, pulando processamento", "WARNING")
                return {
                    'guia': dados['guia'],
                    'status': 'erro',
                    'erro': 'Cartão vazio - processamento pulado',
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
            
            # 1. Preencher número da carteira do beneficiário (17 dígitos)
            log_message("🔍 Preenchendo campo do número da carteira...", "INFO")
            try:
                cartao_formatado = self.formatar_cartao_17_digitos(dados['cartao'])
                
                # Usar JavaScript para preencher o campo conforme sugerido
                javascript_code = f'$("#codigo").val("{cartao_formatado}").trigger("input").trigger("change").trigger("blur");'
                
                # Aguardar o campo estar presente
                wait.until(EC.presence_of_element_located((By.ID, "codigo")))
                
                # Executar JavaScript
                log_message(f"🔧 Executando JavaScript: {javascript_code}", "INFO")
                driver.execute_script(javascript_code)
                log_message(f"✅ Cartão preenchido via JavaScript: {cartao_formatado}", "SUCCESS")
                
                # Verificar se houve erro de carteirinha inválida
                verificacao_erro = self.verificar_erro_carteirinha(driver, wait)
                if verificacao_erro['erro']:
                    log_message(f"❌ Carteirinha inválida: {verificacao_erro['mensagem']}", "ERROR")
                    return {
                        'guia': dados['guia'],
                        'status': 'erro',
                        'erro': verificacao_erro['mensagem'],
                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                
            except Exception as e:
                log_message(f"❌ Erro ao preencher cartão: {e}", "ERROR")
                raise Exception(f"Falha ao preencher número da carteira: {e}")
            
            # 2. Buscar médico solicitante
            log_message("🔍 Iniciando busca do médico solicitante...", "INFO")
            try:
                self.buscar_medico_solicitante(driver, wait, dados['crm'], dados['medico'])
            except Exception as e:
                log_message(f"⚠️ Erro na busca do médico: {e}. Continuando para próximo exame...", "WARNING")
                return {
                    'guia': dados['guia'],
                    'status': 'erro',
                    'erro': f"Erro na busca do médico: {e}",
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
            
            # 3. Preencher campos fixos
            log_message("📝 Preenchendo campos fixos...", "INFO")
            try:
                self.preencher_campos_fixos(driver)
            except Exception as e:
                log_message(f"⚠️ Erro ao preencher campos fixos: {e}", "WARNING")
            
            # 4. Preencher hipótese diagnóstica
            log_message("🔍 Preenchendo hipótese diagnóstica...", "INFO")
            try:
                self.preencher_hipotese_diagnostica(driver, wait, dados['texto'])
            except Exception as e:
                log_message(f"⚠️ Erro ao preencher hipótese diagnóstica: {e}", "WARNING")
            
            # 5. Preencher procedimentos e quantidades
            log_message("📝 Preenchendo procedimentos e quantidades...", "INFO")
            try:
                self.preencher_procedimentos(driver, dados['procedimentos'], dados['qtd'])
            except Exception as e:
                log_message(f"⚠️ Erro ao preencher procedimentos: {e}", "WARNING")
            
            # 6. Autorizar guia e capturar resultado
            log_message("🔄 Autorizando guia...", "INFO")
            try:
                resultado_autorizacao = self.autorizar_guia(driver, wait)
                
                if resultado_autorizacao['sucesso']:
                    # Verificar se é sucesso normal ou enviado para análise
                    if resultado_autorizacao.get('status_especial') == 'analise':
                        log_message(f"⚠️ Guia enviada para análise: {resultado_autorizacao['numero_guia']}", "WARNING")
                        status_resultado = 'analise'
                    else:
                        log_message(f"✅ Guia autorizada com sucesso: {resultado_autorizacao['numero_guia']}", "SUCCESS")
                        status_resultado = 'sucesso'
                    
                    return {
                        'guia': dados['guia'],
                        'status': status_resultado,
                        'numero_guia': resultado_autorizacao['numero_guia'],
                        'mensagem': resultado_autorizacao['mensagem'],
                        'cartao_formatado': cartao_formatado,
                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'status_especial': resultado_autorizacao.get('status_especial')
                    }
                else:
                    log_message(f"❌ Erro na autorização: {resultado_autorizacao['mensagem']}", "ERROR")
                    return {
                        'guia': dados['guia'],
                        'status': 'erro',
                        'erro': resultado_autorizacao['mensagem'],
                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    
            except Exception as e:
                log_message(f"❌ Erro ao autorizar guia: {e}", "ERROR")
                return {
                    'guia': dados['guia'],
                    'status': 'erro',
                    'erro': f"Erro na autorização: {e}",
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
            
        except Exception as e:
            log_message(f"❌ Erro ao processar guia {dados['guia']}: {e}", "ERROR")
            return {
                'guia': dados['guia'],
                'status': 'erro',
                'erro': str(e),
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

    def verificar_guia_ja_processada(self, excel_file, guia):
        """Verifica se uma guia já foi processada anteriormente"""
        try:
            # Ler o arquivo Excel
            df = pd.read_excel(excel_file, header=0)
            
            # Encontrar a coluna GUIA (pode ter qualquer case)
            coluna_guia = None
            for col in df.columns:
                if col.upper().strip() == 'GUIA':
                    coluna_guia = col
                    break
            
            if not coluna_guia:
                log_message(f"⚠️ Coluna GUIA não encontrada no Excel", "WARNING")
                return {'ja_processada': False}
            
            # Encontrar colunas de resultado (qualquer case)
            coluna_numero_guia = None
            coluna_status_guia = None
            
            for col in df.columns:
                col_upper = col.upper().strip()
                if col_upper in ['NUMERO_GUIA', 'NUMEROGUIA', 'NUMERO GUIA']:
                    coluna_numero_guia = col
                elif col_upper in ['STATUS_GUIA', 'STATUSGUIA', 'STATUS GUIA']:
                    coluna_status_guia = col
            
            # Encontrar linha correspondente
            mask = df[coluna_guia].astype(str).str.strip() == str(guia).strip()
            indices = df[mask].index
            
            if len(indices) > 0:
                indice = indices[0]
                
                # Verificar se já tem número da guia e status
                numero_guia = ''
                status_guia = ''
                
                if coluna_numero_guia and coluna_numero_guia in df.columns:
                    numero_guia = df.loc[indice, coluna_numero_guia]
                
                if coluna_status_guia and coluna_status_guia in df.columns:
                    status_guia = df.loc[indice, coluna_status_guia]
                
                # Se tem número da guia e status não está vazio, considera já processada
                if pd.notna(numero_guia) and str(numero_guia).strip() and pd.notna(status_guia) and str(status_guia).strip():
                    return {
                        'ja_processada': True,
                        'numero_guia': str(numero_guia).strip(),
                        'status_guia': str(status_guia).strip()
                    }
            
            return {'ja_processada': False}
            
        except Exception as e:
            log_message(f"⚠️ Erro ao verificar se guia já foi processada: {e}", "WARNING")
            return {'ja_processada': False}

    def esperar_liberacao_guia(self, driver, wait, numero_guia, cancel_flag, max_tentativas=5, tempo_espera=30):
        """Espera a guia ser liberada, consultando o status repetidamente"""
        log_message(f"⏳ Aguardando liberação da guia {numero_guia}...", "INFO")
        for tentativa in range(1, max_tentativas + 1):
            if cancel_flag and cancel_flag.is_set():
                log_message("Execução cancelada pelo usuário durante espera por liberação.", "WARNING")
                return {'sucesso': False, 'status_guia': 'Cancelado', 'numero_guia': numero_guia, 'erro': 'Cancelado pelo usuário'}

            log_message(f"🔍 Tentativa {tentativa}/{max_tentativas}: Consultando status da guia {numero_guia}...", "INFO")
            status_resultado = self.consultar_status_guia(driver, wait, numero_guia)

            if status_resultado.get('sucesso') and status_resultado.get('status_guia', '').strip().lower() == 'liberada':
                log_message(f"✅ Guia {numero_guia} liberada com sucesso!", "SUCCESS")
                return status_resultado
            else:
                log_message(f"ℹ️ Guia {numero_guia} ainda em status: {status_resultado.get('status_guia', 'Erro ao consultar')}. Aguardando {tempo_espera} segundos...", "INFO")
                time.sleep(tempo_espera)

        log_message(f"❌ Guia {numero_guia} não foi liberada após {max_tentativas} tentativas.", "ERROR")
        return {'sucesso': False, 'status_guia': 'Não Liberada', 'numero_guia': numero_guia, 'erro': 'Não liberada após tentativas'}

    def run(self, params: dict):
        username = params.get("unimed_user")
        password = params.get("unimed_pass")
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode")
        excel_file = params.get("excel_file")
        
        # Configurar modo headless na instância
        self.headless_mode = headless_mode
        log_message(f"🔧 Modo headless: {'Ativado' if headless_mode else 'Desativado'}", "INFO")

        # Validar credenciais da Unimed
        if not username or not password:
            messagebox.showerror("Erro", "Credenciais da Unimed são obrigatórias para este módulo.")
            return

        # Validar arquivo Excel
        if not excel_file or not os.path.exists(excel_file):
            messagebox.showerror("Erro", "Arquivo Excel é obrigatório para este módulo.")
            return

        driver = BrowserFactory.create_chrome(headless=headless_mode)
        wait = WebDriverWait(driver, 15)

        try:
            log_message("Iniciando automação de Lançamento de Guia Unimed...", "INFO")

            # Ler dados do Excel
            try:
                dados_excel = self.read_excel_data(excel_file)
                log_message(f"✅ Carregados {len(dados_excel)} registros do Excel", "SUCCESS")
                
                # Exibir amostra dos dados carregados
                if dados_excel:
                    primeiro_registro = dados_excel[0]
                    log_message(f"📋 Exemplo de registro: Guia={primeiro_registro['guia']}, "
                              f"Cartão={primeiro_registro['cartao']}, "
                              f"Médico={primeiro_registro['medico']}", "INFO")
                
            except Exception as e:
                log_message(f"❌ Erro ao ler arquivo Excel: {e}", "ERROR")
                messagebox.showerror("Erro", f"Erro ao ler arquivo Excel:\n{e}")
                return

            # Verificar quais guias já foram processadas
            log_message("🔍 Verificando guias já processadas...", "INFO")
            guias_ja_processadas = []
            guias_para_processar = []
            
            for dados in dados_excel:
                verificacao = self.verificar_guia_ja_processada(excel_file, dados['guia'])
                if verificacao['ja_processada']:
                    guias_ja_processadas.append({
                        'guia': dados['guia'],
                        'numero_guia': verificacao['numero_guia'],
                        'status_guia': verificacao['status_guia'],
                        'status': 'ja_processada'
                    })
                    log_message(f"✅ Guia {dados['guia']} já processada - Número: {verificacao['numero_guia']}, Status: {verificacao['status_guia']}", "SUCCESS")
                else:
                    guias_para_processar.append(dados)
            
            log_message(f"📊 Resumo da verificação:", "INFO")
            log_message(f"   - Guias já processadas: {len(guias_ja_processadas)}", "SUCCESS")
            log_message(f"   - Guias para processar: {len(guias_para_processar)}", "INFO")

            # Processar apenas guias que ainda não foram processadas
            resultados_processamento = []
            
            if guias_para_processar:
                log_message(f"\n🚀 Processando {len(guias_para_processar)} guias na Unimed...", "INFO")
                
                # Fazer login na Unimed apenas se há guias para processar
                self.fazer_login_unimed(driver, wait, username, password)
                
                # Acessar página de procedimentos
                self.acessar_pagina_procedimento(driver)

                # Processar cada registro do Excel
                for i, dados in enumerate(guias_para_processar, 1):
                    if cancel_flag and cancel_flag.is_set():
                        log_message("Execução cancelada pelo usuário.", "WARNING")
                        break
                    
                    try:
                        log_message(f"➡️ Processando registro {i}/{len(guias_para_processar)} - Guia: {dados['guia']}", "INFO")
                        
                        resultado = self.processar_guia_unimed(driver, wait, dados)
                        resultados_processamento.append(resultado)
                        
                        if resultado.get('status') == 'sucesso':
                            log_message(f"✅ Guia {dados['guia']} processada com sucesso - Número: {resultado.get('numero_guia')}", "SUCCESS")
                        else:
                            log_message(f"❌ Erro na guia {dados['guia']}: {resultado.get('erro')}", "ERROR")
                        
                        # Recarregar página para próxima guia (se não for a última)
                        if i < len(guias_para_processar):
                            log_message("🔄 Recarregando página para próxima guia...", "INFO")
                            self.acessar_pagina_procedimento(driver)
                            time.sleep(2)
                        
                    except Exception as e:
                        log_message(f"❌ Erro ao processar guia {dados['guia']}: {e}", "ERROR")
                        resultados_processamento.append({
                            'guia': dados['guia'],
                            'status': 'erro',
                            'erro': str(e)
                        })
            else:
                log_message("ℹ️ Todas as guias já foram processadas, pulando lançamento na Unimed", "INFO")

            # Adicionar guias já processadas aos resultados
            resultados_processamento.extend(guias_ja_processadas)

            # Consultar status das guias criadas (sucesso + análise)
            log_message("\n🔍 Consultando status das guias criadas...", "INFO")
            guias_para_consultar = [r for r in resultados_processamento 
                                   if r.get('status') in ['sucesso', 'analise'] and r.get('numero_guia')]
            
            if guias_para_consultar:
                log_message(f"📋 {len(guias_para_consultar)} guias para consultar status", "INFO")
                
                for i, resultado in enumerate(guias_para_consultar, 1):
                    if cancel_flag and cancel_flag.is_set():
                        log_message("Execução cancelada pelo usuário.", "WARNING")
                        break
                    
                    numero_guia = resultado.get('numero_guia')
                    log_message(f"🔍 Consultando status {i}/{len(guias_para_consultar)} - Guia: {numero_guia}", "INFO")
                    
                    try:
                        status_resultado = self.consultar_status_guia(driver, wait, numero_guia)
                        
                        # Atualizar o resultado com o status da guia
                        resultado['status_guia'] = status_resultado.get('status_guia', 'Erro ao consultar')
                        
                        if status_resultado.get('sucesso'):
                            log_message(f"✅ Status da guia {numero_guia}: {status_resultado.get('status_guia')}", "SUCCESS")
                           
                            # Se a guia não estiver liberada, esperar pela liberação
                            if status_resultado.get('status_guia', '').strip().lower() != 'liberada':
                                log_message(f"⏳ Guia {numero_guia} não está liberada. Aguardando...", "INFO")
                                liberacao_resultado = self.esperar_liberacao_guia(driver, wait, numero_guia, cancel_flag)
                                if not liberacao_resultado.get('sucesso'):
                                    log_message(f"❌ Guia {numero_guia} não foi liberada após tentativas. Número preservado no Excel para consulta.", "ERROR")
                                    resultado['status'] = 'erro'
                                    resultado['erro'] = liberacao_resultado.get('erro', 'Guia não liberada')
                                    resultado['status_guia'] = liberacao_resultado.get('status_guia', 'Não Liberada')
                                    # numero_guia já está em resultado, será preservado no Excel
                                    continue # Pular para a próxima guia, pois esta não foi liberada
                                else:
                                    resultado['status_guia'] = liberacao_resultado.get('status_guia')
                                    log_message(f"✅ Guia {numero_guia} agora está: {resultado['status_guia']}", "SUCCESS")
                        else:
                            log_message(f"⚠️ Erro ao consultar guia {numero_guia}: {status_resultado.get('erro')}", "WARNING")
                        
                        # Aguardar entre consultas
                        time.sleep(2)
                        
                    except Exception as e:
                        log_message(f"❌ Erro ao consultar status da guia {numero_guia}: {e}", "ERROR")
                        resultado['status_guia'] = 'Erro ao consultar'
            else:
                log_message("ℹ️ Nenhuma guia nova foi criada para consultar status", "INFO")

            # Acessar PathoWeb para abrir exames das guias (todas as que têm número)
            log_message("\n🌐 Acessando PathoWeb para abrir exames...", "INFO")
            
            # Obter credenciais do PathoWeb dos parâmetros (usar username e password padrão)
            pathoweb_user = params.get("username")  # Mudança: usar username em vez de pathoweb_user
            pathoweb_pass = params.get("password")  # Mudança: usar password em vez de pathoweb_pass
            
            # Filtrar guias que realmente podem ser abertas (com número e, se houver, status liberada)
            guias_para_abrir = [r for r in resultados_processamento 
                                   if r.get('numero_guia') and (not r.get('status_guia') or str(r.get('status_guia')).strip().lower() == 'liberada')]
            
            pathoweb_sucessos = 0
            
            if not guias_para_abrir:
                log_message("ℹ️ Nenhuma guia liberada com número para abrir no PathoWeb. Pulando acesso ao PathoWeb.", "INFO")
            else:
                if pathoweb_user and pathoweb_pass:
                    # Fazer login no PathoWeb
                    if self.fazer_login_pathoweb(driver, wait, pathoweb_user, pathoweb_pass):
                        log_message("✅ Login no PathoWeb realizado com sucesso", "SUCCESS")
                        
                        log_message(f"📋 {len(guias_para_abrir)} guias para abrir no PathoWeb", "INFO")
                        
                        for i, resultado in enumerate(guias_para_abrir, 1):
                            if cancel_flag and cancel_flag.is_set():
                                log_message("Execução cancelada pelo usuário.", "WARNING")
                                break
                            
                            numero_guia_unimed = resultado.get('numero_guia')
                            guia_original = resultado.get('guia')  # Número da guia original da coluna A
                            status_guia = resultado.get('status_guia', 'Status não consultado')
                            log_message(f"🔍 Abrindo exame {i}/{len(guias_para_abrir)} - Guia Original: {guia_original} (Unimed: {numero_guia_unimed}, Status: {status_guia})", "INFO")
                            
                            try:
                                sucesso_abertura = self.abrir_exame_pathoweb(driver, wait, guia_original, numero_guia_unimed)
                                
                                if sucesso_abertura:
                                    pathoweb_sucessos += 1
                                    log_message(f"✅ Exame {guia_original} aberto com sucesso no PathoWeb", "SUCCESS")
                                else:
                                    log_message(f"❌ Erro ao abrir exame {guia_original} no PathoWeb", "ERROR")
                                
                                # Navegar de volta para página de busca se não for o último exame
                                if i < len(guias_para_abrir):
                                    log_message(f"🔄 Retornando para página de busca para próximo exame ({i+1}/{len(guias_para_abrir)})...", "INFO")
                                    try:
                                        driver.get("https://dap.pathoweb.com.br/moduloFaturamento/index")
                                        time.sleep(2)
                                        
                                        # Clicar em "Preparar exames para fatura" novamente
                                        try:
                                            preparar_btn = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                                                "a.btn.btn-danger.chamadaAjax.setupAjax[data-url='/moduloFaturamento/preFaturamento']",
                                                condition="presence")
                                            self.click_element(driver, preparar_btn, "botão 'Preparar exames' (reload)")
                                        except Exception:
                                            preparar_btn = self.wait_for_element(driver, wait, By.XPATH,
                                                "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]",
                                                condition="presence")
                                            self.click_element(driver, preparar_btn, "botão 'Preparar exames' (reload alt)")
                                        
                                        # Aguardar spinner se existir
                                        try:
                                            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                                            WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                                        except Exception:
                                            time.sleep(2)
                                        
                                        log_message("✅ Página de busca recarregada", "SUCCESS")
                                    except Exception as e:
                                        log_message(f"⚠️ Erro ao recarregar página de busca: {e}", "WARNING")
                                else:
                                    log_message(f"✅ Último exame processado ({i}/{len(guias_para_abrir)})", "SUCCESS")
                                
                                time.sleep(2)
                                
                            except Exception as e:
                                log_message(f"❌ Erro ao abrir exame {guia_original}: {e}", "ERROR")
                    else:
                        log_message("❌ Falha no login do PathoWeb", "ERROR")
                else:
                    log_message("⚠️ Credenciais do PathoWeb não fornecidas, pulando acesso ao PathoWeb", "WARNING")

            # Salvar resultados no Excel
            try:
                arquivo_resultados = self.salvar_resultados_excel(excel_file, resultados_processamento)
                if arquivo_resultados:
                    log_message(f"📊 Resultados salvos em: {arquivo_resultados}", "SUCCESS")
                else:
                    log_message("⚠️ Não foi possível salvar os resultados no Excel", "WARNING")
            except Exception as e:
                log_message(f"❌ Erro ao salvar resultados: {e}", "ERROR")

            # Resumo final
            total = len(resultados_processamento)
            sucessos = sum(1 for r in resultados_processamento if r.get('status') == 'sucesso')
            analises = sum(1 for r in resultados_processamento if r.get('status') == 'analise')
            ja_processadas = sum(1 for r in resultados_processamento if r.get('status') == 'ja_processada')
            erros = sum(1 for r in resultados_processamento if r.get('status') == 'erro')
            guias_com_status = sum(1 for r in resultados_processamento if r.get('status_guia'))

            log_message(f"\n📊 Resumo do processamento:", "INFO")
            log_message(f"Total de registros: {total}", "INFO")
            log_message(f"Sucessos (aprovadas): {sucessos}", "SUCCESS" if sucessos > 0 else "INFO")
            log_message(f"Enviadas para análise: {analises}", "WARNING" if analises > 0 else "INFO")
            log_message(f"Já processadas: {ja_processadas}", "SUCCESS" if ja_processadas > 0 else "INFO")
            log_message(f"Erros: {erros}", "ERROR" if erros > 0 else "INFO")
            log_message(f"Status consultados: {guias_com_status}", "INFO")

            # Lista dos números das guias geradas (sucesso + análise)
            guias_com_numero = [r.get('numero_guia') for r in resultados_processamento if r.get('numero_guia')]
            guias_sucesso_direto = [r.get('numero_guia') for r in resultados_processamento if r.get('status') == 'sucesso' and r.get('numero_guia')]
            guias_analise = [r.get('numero_guia') for r in resultados_processamento if r.get('status') == 'analise' and r.get('numero_guia')]
            
            if guias_com_numero:
                log_message(f"🎯 Guias com número gerado: {', '.join(guias_com_numero)}", "SUCCESS")
                if guias_sucesso_direto:
                    log_message(f"✅ Aprovadas diretamente: {', '.join(guias_sucesso_direto)}", "SUCCESS")
                if guias_analise:
                    log_message(f"⚠️ Enviadas para análise: {', '.join(guias_analise)}", "WARNING")

            mensagem_final = f"✅ Processamento finalizado!\n\n" \
                           f"Total de registros: {total}\n" \
                           f"Sucessos (aprovadas): {sucessos}\n" \
                           f"Enviadas para análise: {analises}\n" \
                           f"Já processadas: {ja_processadas}\n" \
                           f"Erros: {erros}\n" \
                           f"Status consultados: {guias_com_status}"
            
            if arquivo_resultados:
                mensagem_final += f"\n\n📊 Resultados salvos em:\n{arquivo_resultados}"

            messagebox.showinfo("Processamento Concluído", mensagem_final)

            # Retorno conciso indicando sucesso geral e resultados detalhados
            return {
                'sucesso': (sucessos + analises > 0) or (pathoweb_sucessos > 0),
                'unimed_sucessos': sucessos,
                'unimed_analises': analises,
                'pathoweb_sucessos': pathoweb_sucessos,
                'resultados': resultados_processamento,
                'arquivo_resultados': arquivo_resultados
            }

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            # Aguardar antes de fechar para permitir visualização dos resultados
            if not headless_mode:
                input("Pressione Enter para fechar o navegador...")
            driver.quit()

    def marcar_exames_como_conferidos(self, driver, wait):
        """Marca a coluna 'Conferido' para todos os exames na tabela"""
        try:
            log_message("📝 Marcando exames como 'Conferido' na tabela...", "INFO")
            time.sleep(2)

            # Re-localizar a tabela sempre antes de processar para evitar elementos stale
            def obter_linhas():
                return driver.find_elements(By.CSS_SELECTOR, "#tabelaPreFaturamentoTbody tr")
            
            linhas_iniciais = obter_linhas()
            if not linhas_iniciais:
                log_message("⚠️ Nenhuma linha encontrada na tabela de pré-faturamento", "WARNING")
                return

            total_linhas = len(linhas_iniciais)
            log_message(f"📋 Total de linhas encontradas: {total_linhas}", "INFO")
            
            # Processar cada linha por índice (re-localizando elementos a cada iteração)
            linhas_processadas = 0
            
            for idx in range(total_linhas):
                try:
                    log_message(f"🔄 Processando linha {idx + 1}/{total_linhas}...", "INFO")
                    
                    # SEMPRE re-localizar elementos para evitar stale elements
                    # Aguardar spinner desaparecer antes de re-localizar
                    try:
                        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.ID, "spinner")))
                        log_message(f"⏳ Aguardando spinner desaparecer antes de processar linha {idx + 1}...", "INFO")
                        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                    except Exception:
                        pass
                    
                    time.sleep(0.5)  # Pequena pausa para estabilidade
                    
                    # Re-localizar todas as linhas
                    linhas_atuais = obter_linhas()
                    if idx >= len(linhas_atuais):
                        log_message(f"⚠️ Linha {idx + 1} não existe mais na tabela (total atual: {len(linhas_atuais)})", "WARNING")
                        continue
                    
                    linha = linhas_atuais[idx]
                    
                    # Re-localizar células dentro da linha atual
                    celulas = linha.find_elements(By.CSS_SELECTOR, "td")
                    if len(celulas) < 2:
                        log_message(f"⚠️ Linha {idx + 1}: células insuficientes ({len(celulas)})", "WARNING")
                        continue

                    # Segunda coluna é a de 'Conferido'
                    cel_conferido = celulas[1]

                    # Verificar se já está marcado como 'Conferido'
                    try:
                        ancora = cel_conferido.find_element(By.CSS_SELECTOR, "a.table-editable-ancora")
                        texto_ancora = (ancora.text or "").strip().lower()
                        if texto_ancora == "conferido":
                            log_message(f"✅ Linha {idx + 1}: já está 'Conferido'", "SUCCESS")
                            linhas_processadas += 1
                            continue
                    except Exception:
                        # Se não encontrar âncora, tentar processar mesmo assim
                        log_message(f"ℹ️ Linha {idx + 1}: âncora não encontrada, tentando processar", "INFO")

                    # Tentar abrir o editor clicando na âncora
                    clicou_ancora = False
                    for tentativa in range(3):  # Até 3 tentativas para clicar
                        try:
                            # Re-localizar âncora para evitar stale
                            ancora = cel_conferido.find_element(By.CSS_SELECTOR, "a.table-editable-ancora")
                            
                            # Em modo headless, não fazer scroll (pode causar problemas)
                            if not self.headless_mode:
                                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", ancora)
                                time.sleep(0.3)
                            
                            # Aguardar spinner invisível
                            try:
                                WebDriverWait(driver, 2).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                            except Exception:
                                pass
                            
                            # Usar método robusto de clique
                            self.click_element(driver, ancora, f"âncora linha {idx + 1}")
                            time.sleep(0.5)
                            clicou_ancora = True
                            log_message(f"✅ Linha {idx + 1}: clicou na âncora (tentativa {tentativa + 1})", "INFO")
                            break
                            
                        except Exception as e:
                            log_message(f"⚠️ Linha {idx + 1}: erro ao clicar na âncora (tentativa {tentativa + 1}): {e}", "WARNING")
                            if tentativa < 2:
                                # Aguardar spinner e tentar novamente
                                try:
                                    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                                    time.sleep(0.5)
                                except Exception:
                                    time.sleep(1)
                    
                    if not clicou_ancora:
                        log_message(f"❌ Linha {idx + 1}: não conseguiu clicar na âncora após 3 tentativas", "ERROR")
                        continue

                    # Selecionar 'Conferido' no select
                    selecionou = False
                    for tentativa in range(3):  # Até 3 tentativas para selecionar
                        try:
                            # Re-localizar a célula e o select
                            linhas_temp = obter_linhas()
                            if idx < len(linhas_temp):
                                cel_conferido_temp = linhas_temp[idx].find_elements(By.CSS_SELECTOR, "td")[1]
                                select_el = cel_conferido_temp.find_element(By.CSS_SELECTOR, "select[name='faturamentoConferido']")
                                
                                # Usar JavaScript para garantir a seleção
                                driver.execute_script("""
                                    var s = arguments[0];
                                    $(s).val('Conferido').trigger('change').trigger('blur');
                                """, select_el)
                                
                                log_message(f"✅ Linha {idx + 1}: selecionou 'Conferido' (tentativa {tentativa + 1})", "SUCCESS")
                                selecionou = True
                                linhas_processadas += 1
                                break
                        except Exception as e:
                            log_message(f"⚠️ Linha {idx + 1}: erro ao selecionar 'Conferido' (tentativa {tentativa + 1}): {e}", "WARNING")
                            if tentativa < 2:
                                time.sleep(0.5)
                    
                    if not selecionou:
                        log_message(f"❌ Linha {idx + 1}: não conseguiu selecionar 'Conferido' após 3 tentativas", "ERROR")
                        continue

                    # Aguardar processamento (spinner)
                    try:
                        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                        log_message(f"🔄 Linha {idx + 1}: processando alteração (spinner detectado)...", "INFO")
                        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                        log_message(f"✅ Linha {idx + 1}: processamento concluído", "SUCCESS")
                    except Exception:
                        # Sem spinner; pequena pausa
                        time.sleep(0.5)
                        log_message(f"ℹ️ Linha {idx + 1}: sem spinner, aguardando estabilização", "INFO")

                except Exception as e:
                    log_message(f"❌ Erro crítico ao processar linha {idx + 1}: {e}", "ERROR")
                    # Continuar para próxima linha mesmo com erro
                    continue

            log_message(f"✅ Processamento concluído: {linhas_processadas}/{total_linhas} linhas marcadas como 'Conferido'", "SUCCESS")
            
            # Aguardar processamento final (especialmente importante quando há apenas 1 exame)
            log_message("⏳ Aguardando processamento final antes de continuar...", "INFO")
            try:
                # Tentar detectar se há spinner ativo
                WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "spinner")))
                log_message("🔄 Spinner final detectado, aguardando conclusão...", "INFO")
                WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                log_message("✅ Spinner final concluído", "SUCCESS")
            except Exception:
                # Se não houver spinner, aguardar tempo fixo para garantir
                log_message("ℹ️ Spinner não detectado, aguardando tempo de segurança...", "INFO")
                time.sleep(2)
            
            # Verificação final
            log_message("📋 Realizando verificação final...", "INFO")
            time.sleep(1)
            linhas_finais = obter_linhas()
            conferidos_final = 0
            for linha_final in linhas_finais:
                try:
                    celulas_final = linha_final.find_elements(By.CSS_SELECTOR, "td")
                    if len(celulas_final) >= 2:
                        ancora_final = celulas_final[1].find_element(By.CSS_SELECTOR, "a.table-editable-ancora")
                        if (ancora_final.text or "").strip().lower() == "conferido":
                            conferidos_final += 1
                except Exception:
                    pass
            
            log_message(f"📊 Verificação final: {conferidos_final}/{len(linhas_finais)} exames estão marcados como 'Conferido'", "INFO")
            
            # Tempo adicional de segurança antes de fechar/prosseguir
            if conferidos_final == total_linhas and total_linhas > 0:
                log_message("✅ Todos os exames foram marcados com sucesso, aguardando estabilização...", "SUCCESS")
                time.sleep(2)
            elif conferidos_final < total_linhas:
                log_message(f"⚠️ Alguns exames podem não ter sido marcados ({conferidos_final}/{total_linhas}), aguardando tempo adicional...", "WARNING")
                time.sleep(3)
            
        except Exception as e:
            log_message(f"❌ Erro ao marcar exames como 'Conferido': {e}", "ERROR")


def run(params: dict):
    module = LancamentoGuiaUnimedModule()
    module.run(params) 