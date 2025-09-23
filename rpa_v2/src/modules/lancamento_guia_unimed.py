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

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

class LancamentoGuiaUnimedModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Lançamento Guia Unimed")

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
                    data_list.append({
                        'guia': str(row['GUIA']).strip(),
                        'cartao': str(row['CARTAO']).strip() if pd.notna(row['CARTAO']) else '',
                        'medico': str(row['MEDICO']).strip() if pd.notna(row['MEDICO']) else '',
                        'crm': str(row['CRM']).strip() if pd.notna(row['CRM']) else '',
                        'procedimentos': str(row['PROCEDIMENTOS']).strip() if pd.notna(row['PROCEDIMENTOS']) else '',
                        'qtd': str(row['QTD']).strip() if pd.notna(row['QTD']) else '',
                        'texto': str(row['TEXTO']).strip() if pd.notna(row['TEXTO']) else ''
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
            (r'B[ií]ópsia\s+G[aá]strica', 'Biópsias Gástricas'),
            (r'B[ií]ópsia\s+de\s+Pr[oó]stata', 'Biópsias de Próstata'),
            (r'B[ií]ópsia\s+de\s+Reto', 'Biópsias de Reto'),
            (r'B[ií]ópsia\s+G[aá]strica', 'Biópsias Gástricas'),
            (r'B[ií]ópsia\s+Pr[oó]stata', 'Biópsias de Próstata'),
            (r'B[ií]ópsia\s+Reto', 'Biópsias de Reto'),
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

    def formatar_cartao_17_digitos(self, cartao):
        """Formata o número do cartão para ter 17 dígitos, adicionando zeros antes se necessário"""
        cartao_limpo = str(cartao).strip()
        
        # Remover apóstrofe do Excel e outros caracteres especiais, manter apenas números e letras
        cartao_sem_apostrofe = cartao_limpo.lstrip("'")  # Remove apóstrofe do início
        cartao_sem_espacos = ''.join(cartao_sem_apostrofe.split())
        
        if len(cartao_sem_espacos) < 17:
            # Adicionar zeros à esquerda para completar 17 dígitos
            zeros_necessarios = 17 - len(cartao_sem_espacos)
            cartao_formatado = "0" * zeros_necessarios + cartao_sem_espacos
            log_message(f"📋 Cartão formatado: '{cartao_limpo}' → {cartao_formatado} (17 dígitos)", "INFO")
            return cartao_formatado
        elif len(cartao_sem_espacos) == 17:
            log_message(f"📋 Cartão já tem 17 dígitos: {cartao_sem_espacos}", "INFO")
            return cartao_sem_espacos
        else:
            log_message(f"⚠️ Cartão com mais de 17 dígitos ({len(cartao_sem_espacos)}): {cartao_sem_espacos}", "WARNING")
            return cartao_sem_espacos

    def extrair_apenas_numeros(self, crm):
        """Extrai apenas os números do CRM, removendo letras"""
        apenas_numeros = re.sub(r'[^0-9]', '', str(crm))
        log_message(f"📋 CRM formatado: {crm} → {apenas_numeros}", "INFO")
        return apenas_numeros

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
            
            # 4. Preencher campo do conselho com números do CRM
            log_message(f"📝 Preenchendo campo conselho com: {crm_numeros}", "INFO")
            campo_conselho = wait.until(EC.presence_of_element_located((By.ID, "conselho")))
            campo_conselho.clear()
            campo_conselho.send_keys(crm_numeros)
            
            # 4.1. Preencher campo do nome do médico para filtrar melhor
            log_message(f"📝 Preenchendo campo nome com: {nome_medico}", "INFO")
            try:
                campo_nome = wait.until(EC.presence_of_element_located((By.ID, "nome")))
                campo_nome.clear()
                campo_nome.send_keys(nome_medico.upper() if nome_medico else "")
                log_message(f"✅ Nome do médico preenchido: {nome_medico}", "SUCCESS")
            except Exception as e:
                log_message(f"⚠️ Erro ao preencher nome do médico: {e}", "WARNING")
            
            # 5. Clicar no botão localizar
            log_message("🔍 Clicando em localizar...", "INFO")
            botao_localizar = wait.until(EC.element_to_be_clickable((By.ID, "localizar")))
            botao_localizar.click()
            
            # 6. Aguardar tabela carregar
            time.sleep(3)
            
            # 7. Verificar se tabela foi carregada
            try:
                tabela = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.table-hover tbody")))
                linhas = tabela.find_elements(By.CSS_SELECTOR, "tr")
                
                # Verificar se há resultados (primeira linha é cabeçalho)
                if len(linhas) <= 1:
                    log_message(f"⚠️ Nenhum médico encontrado para CRM: {crm}", "WARNING")
                    raise Exception(f"Médico não encontrado para CRM: {crm}")
                
                # 8. Selecionar médico (deve ser único com CRM + nome filtrados)
                medico_encontrado = False
                
                # Com CRM e nome preenchidos, deve retornar apenas um resultado
                for linha in linhas[1:]:  # Pula o cabeçalho
                    try:
                        colunas = linha.find_elements(By.TAG_NAME, "td")
                        if len(colunas) >= 3:
                            codigo_medico = colunas[0].text.strip()
                            nome_na_tabela = colunas[1].text.strip()
                            documento = colunas[2].text.strip()
                            
                            log_message(f"📋 Médico encontrado: {nome_na_tabela} - {documento}", "INFO")
                            
                            # Clicar na linha do médico (deve ser único)
                            linha.click()
                            log_message(f"✅ Médico selecionado: {nome_na_tabela} - {documento}", "SUCCESS")
                            medico_encontrado = True
                            break
                                    
                    except Exception as e:
                        log_message(f"⚠️ Erro ao processar linha da tabela: {e}", "WARNING")
                        continue
                
                if not medico_encontrado:
                    log_message(f"⚠️ Médico não encontrado na tabela para CRM: {crm}", "WARNING")
                    raise Exception(f"Médico não encontrado na tabela para CRM: {crm}")
                
                # 9. Aguardar popup fechar automaticamente e voltar para janela original
                time.sleep(3)
                
                # O popup fecha automaticamente, então só precisamos voltar para janela original
                driver.switch_to.window(janela_original)
                log_message("✅ Médico selecionado, voltou para janela principal", "SUCCESS")
                
            except Exception as e:
                log_message(f"❌ Erro ao processar tabela de médicos: {e}", "ERROR")
                # Tentar voltar para janela original em caso de erro
                try:
                    driver.switch_to.window(janela_original)
                    log_message("🔄 Voltou para janela principal após erro", "INFO")
                except:
                    pass
                raise Exception(f"Falha na busca do médico: {e}")
                
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

    def autorizar_guia(self, driver, wait):
        """Clica no botão Autorizar e captura o resultado"""
        try:
            log_message("🔄 Clicando no botão Autorizar...", "INFO")
            
            # Clicar no botão autorizar
            botao_autorizar = wait.until(EC.element_to_be_clickable((By.ID, "autorizar")))
            botao_autorizar.click()
            
            # Aguardar modal aparecer
            time.sleep(3)
            
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
                match_guia = re.search(r'Guia\s+(\d+)\s+gerada com sucesso', texto_modal)
                
                if match_guia:
                    numero_guia = match_guia.group(1)
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
                else:
                    # Se não encontrou padrão de sucesso, é erro
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
            url = "https://pathoweb.com.br/login/auth"
            driver.get(url)
            
            # Preencher credenciais
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
                    log_message("🔄 Navegação direta para módulo realizada", "INFO")

            elif "moduloFaturamento" in current_url:
                log_message("✅ Já está no módulo de faturamento - pulando navegação", "SUCCESS")
            else:
                log_message(f"⚠️ URL inesperada detectada: {current_url}", "WARNING")
                driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                time.sleep(2)
                log_message("🔄 Navegação direta para módulo realizada (fallback)", "INFO")

            # Fechar modal se aparecer
            try:
                modal_close_button = driver.find_element(By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button")
                if modal_close_button.is_displayed():
                    modal_close_button.click()
                    time.sleep(1)
            except Exception:
                pass

            # Acessar explicitamente a página do módulo de faturamento
            log_message("Acessando módulo de faturamento via URL...", "INFO")
            driver.get("https://pathoweb.com.br/moduloFaturamento/index")

            # Clicar no botão "Preparar exames para fatura"
            log_message("Clicando em 'Preparar exames para fatura'...", "INFO")
            try:
                preparar_btn = wait.until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "a.btn.btn-danger.chamadaAjax.setupAjax[data-url='/moduloFaturamento/preFaturamento']"
                )))
                preparar_btn.click()
            except Exception:
                preparar_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH,
                    "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]"
                )))
                preparar_btn.click()

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

    def preencher_campos_exame(self, driver, numero_guia_unimed):
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
            pesquisar_btn = wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento")))
            pesquisar_btn.click()
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
                    checkbox.click()
                    log_message("✅ Checkbox do exame marcado", "SUCCESS")
                else:
                    log_message("ℹ️ Checkbox já estava marcado", "INFO")
                
                time.sleep(1)
                
                # Procurar e clicar no botão "Abrir exame"
                log_message("🔍 Procurando botão 'Abrir exame'...", "INFO")
                
                abrir_btn = wait.until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR, 
                    "a.btn.btn-sm.btn-primary.chamadaAjax.toogleInicial.setupAjax[data-url='/moduloFaturamento/abrirExameCorrecao']"
                )))
                log_message("✅ Botão 'Abrir exame' encontrado", "SUCCESS")
                
                # Clicar no botão
                abrir_btn.click()
                log_message("✅ Clique no botão 'Abrir exame' realizado", "SUCCESS")
                
                # Aguardar o modal aparecer
                log_message("⏳ Aguardando modal do exame abrir...", "INFO")
                time.sleep(3)
                
                # Verificar se o modal foi aberto
                try:
                    modal = wait.until(EC.presence_of_element_located((By.ID, "myModal")))
                    if modal.is_displayed():
                        log_message("✅ Modal do exame aberto com sucesso", "SUCCESS")
                        
                        # Preencher campos do exame se número da guia Unimed foi fornecido
                        if numero_guia_unimed:
                            log_message("📝 Preenchendo campos do exame...", "INFO")
                            self.preencher_campos_exame(driver, numero_guia_unimed)
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

    def salvar_resultados_excel(self, excel_file, resultados):
        """Salva os resultados no arquivo Excel original"""
        try:
            log_message("💾 Salvando resultados no Excel...", "INFO")
            log_message(f"📁 Arquivo de destino: {excel_file}", "INFO")
            
            # Ler o arquivo Excel original
            df = pd.read_excel(excel_file, header=0)
            log_message(f"📋 Colunas originais: {list(df.columns)}", "INFO")
            
            # Ajustar nomes das colunas
            df.columns = df.columns.str.upper().str.strip()
            log_message(f"📋 Colunas após ajuste: {list(df.columns)}", "INFO")
            
            # Adicionar colunas de resultado se não existirem
            colunas_adicionadas = []
            if 'NUMERO_GUIA' not in df.columns:
                df['NUMERO_GUIA'] = ''
                colunas_adicionadas.append('NUMERO_GUIA')
            if 'STATUS_PROCESSAMENTO' not in df.columns:
                df['STATUS_PROCESSAMENTO'] = ''
                colunas_adicionadas.append('STATUS_PROCESSAMENTO')
            if 'STATUS_GUIA' not in df.columns:
                df['STATUS_GUIA'] = ''
                colunas_adicionadas.append('STATUS_GUIA')
            if 'MENSAGEM_ERRO' not in df.columns:
                df['MENSAGEM_ERRO'] = ''
                colunas_adicionadas.append('MENSAGEM_ERRO')
            if 'DATA_PROCESSAMENTO' not in df.columns:
                df['DATA_PROCESSAMENTO'] = ''
                colunas_adicionadas.append('DATA_PROCESSAMENTO')
            
            if colunas_adicionadas:
                log_message(f"✅ Colunas adicionadas: {colunas_adicionadas}", "SUCCESS")
            else:
                log_message("ℹ️ Todas as colunas já existiam", "INFO")
            
            log_message(f"📋 Colunas finais: {list(df.columns)}", "INFO")
            
            # Atualizar resultados
            for resultado in resultados:
                guia = resultado.get('guia')
                log_message(f"📝 Processando resultado para guia: {guia}", "INFO")
                
                # Encontrar linha correspondente
                mask = df['GUIA'].astype(str).str.strip() == str(guia).strip()
                indices = df[mask].index
                
                if len(indices) > 0:
                    indice = indices[0]
                    log_message(f"✅ Linha encontrada no índice: {indice}", "SUCCESS")
                    
                    # Atualizar dados
                    df.loc[indice, 'DATA_PROCESSAMENTO'] = resultado.get('timestamp', '')
                    df.loc[indice, 'STATUS_GUIA'] = resultado.get('status_guia', '')
                    
                    if resultado.get('status') == 'sucesso':
                        df.loc[indice, 'NUMERO_GUIA'] = resultado.get('numero_guia', '')
                        df.loc[indice, 'STATUS_PROCESSAMENTO'] = 'SUCESSO'
                        df.loc[indice, 'MENSAGEM_ERRO'] = ''
                    else:
                        df.loc[indice, 'NUMERO_GUIA'] = ''
                        df.loc[indice, 'STATUS_PROCESSAMENTO'] = 'ERRO'
                        df.loc[indice, 'MENSAGEM_ERRO'] = resultado.get('erro', resultado.get('mensagem', ''))
                    
                    log_message(f"📝 Resultado salvo para guia {guia}: {df.loc[indice, 'STATUS_PROCESSAMENTO']}", "INFO")
                    log_message(f"   - NUMERO_GUIA: {df.loc[indice, 'NUMERO_GUIA']}", "INFO")
                    log_message(f"   - STATUS_GUIA: {df.loc[indice, 'STATUS_GUIA']}", "INFO")
                else:
                    log_message(f"⚠️ Linha não encontrada para guia: {guia}", "WARNING")
            
            # Estratégia 1: Tentar salvar no arquivo original (sobrescrever)
            try:
                log_message("💾 Tentando salvar no arquivo original...", "INFO")
                df.to_excel(excel_file, index=False)
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
                df.to_excel(nome_arquivo_separado, index=False)
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
                
                # Aguardar um pouco após preencher para ver o resultado
                time.sleep(2)
                
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
                    log_message(f"✅ Guia autorizada com sucesso: {resultado_autorizacao['numero_guia']}", "SUCCESS")
                    return {
                        'guia': dados['guia'],
                        'status': 'sucesso',
                        'numero_guia': resultado_autorizacao['numero_guia'],
                        'mensagem': resultado_autorizacao['mensagem'],
                        'cartao_formatado': cartao_formatado,
                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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
            df.columns = df.columns.str.upper().str.strip()
            
            # Encontrar linha correspondente
            mask = df['GUIA'].astype(str).str.strip() == str(guia).strip()
            indices = df[mask].index
            
            if len(indices) > 0:
                indice = indices[0]
                
                # Verificar se já tem número da guia e status
                numero_guia = df.loc[indice, 'NUMERO_GUIA'] if 'NUMERO_GUIA' in df.columns else ''
                status_guia = df.loc[indice, 'STATUS_GUIA'] if 'STATUS_GUIA' in df.columns else ''
                
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

    def run(self, params: dict):
        username = params.get("unimed_user")
        password = params.get("unimed_pass")
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode")
        excel_file = params.get("excel_file")

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

            # Consultar status das guias criadas com sucesso (apenas as novas)
            log_message("\n🔍 Consultando status das guias criadas...", "INFO")
            guias_para_consultar = [r for r in resultados_processamento 
                                   if r.get('status') == 'sucesso' and r.get('numero_guia')]
            
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
            
            if pathoweb_user and pathoweb_pass:
                # Fazer login no PathoWeb
                if self.fazer_login_pathoweb(driver, wait, pathoweb_user, pathoweb_pass):
                    log_message("✅ Login no PathoWeb realizado com sucesso", "SUCCESS")
                    
                    # Abrir exames de todas as guias que têm número (novas e já processadas)
                    guias_para_abrir = [r for r in resultados_processamento 
                                       if r.get('numero_guia')]
                    
                    if guias_para_abrir:
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
                                # Usar a guia original (da coluna A) para buscar no PathoWeb
                                # Passar também o número da guia Unimed para preencher os campos
                                sucesso_abertura = self.abrir_exame_pathoweb(driver, wait, guia_original, numero_guia_unimed)
                                
                                if sucesso_abertura:
                                    log_message(f"✅ Exame {guia_original} aberto com sucesso no PathoWeb", "SUCCESS")
                                    
                                    # Aguardar um pouco para o usuário visualizar
                                    if not headless_mode:
                                        input(f"Pressione Enter para continuar para o próximo exame (Guia Original: {guia_original})...")
                                else:
                                    log_message(f"❌ Erro ao abrir exame {guia_original} no PathoWeb", "ERROR")
                                
                                # Aguardar entre aberturas
                                time.sleep(2)
                                
                            except Exception as e:
                                log_message(f"❌ Erro ao abrir exame {guia_original}: {e}", "ERROR")
                    else:
                        log_message("ℹ️ Nenhuma guia com número encontrada para abrir no PathoWeb", "INFO")
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
            ja_processadas = sum(1 for r in resultados_processamento if r.get('status') == 'ja_processada')
            erros = sum(1 for r in resultados_processamento if r.get('status') == 'erro')
            guias_com_status = sum(1 for r in resultados_processamento if r.get('status_guia'))

            log_message(f"\n📊 Resumo do processamento:", "INFO")
            log_message(f"Total de registros: {total}", "INFO")
            log_message(f"Sucessos (novos): {sucessos}", "SUCCESS" if sucessos > 0 else "INFO")
            log_message(f"Já processadas: {ja_processadas}", "SUCCESS" if ja_processadas > 0 else "INFO")
            log_message(f"Erros: {erros}", "ERROR" if erros > 0 else "INFO")
            log_message(f"Status consultados: {guias_com_status}", "INFO")

            # Lista dos números das guias geradas com sucesso
            guias_sucesso = [r.get('numero_guia') for r in resultados_processamento if r.get('numero_guia')]
            if guias_sucesso:
                log_message(f"🎯 Guias disponíveis: {', '.join(guias_sucesso)}", "SUCCESS")

            mensagem_final = f"✅ Processamento finalizado!\n\n" \
                           f"Total de registros: {total}\n" \
                           f"Sucessos (novos): {sucessos}\n" \
                           f"Já processadas: {ja_processadas}\n" \
                           f"Erros: {erros}\n" \
                           f"Status consultados: {guias_com_status}"
            
            if arquivo_resultados:
                mensagem_final += f"\n\n📊 Resultados salvos em:\n{arquivo_resultados}"

            messagebox.showinfo("Processamento Concluído", mensagem_final)

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            # Aguardar antes de fechar para permitir visualização dos resultados
            if not headless_mode:
                input("Pressione Enter para fechar o navegador...")
            driver.quit()


def run(params: dict):
    module = LancamentoGuiaUnimedModule()
    module.run(params) 