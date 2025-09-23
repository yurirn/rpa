import os
import json
import numpy as np
from paddleocr import PaddleOCR
from pdf2image import convert_from_path
from PIL import Image
from datetime import datetime
import re
from typing import Dict, Optional

# Configuração dos diretórios
EXAMES_DIR = 'exames/'
RESULTADOS_DIR = 'resultados_ocr/'
os.makedirs(RESULTADOS_DIR, exist_ok=True)

# Inicialização do PaddleOCR otimizada
ocr = PaddleOCR(
    use_angle_cls=True,
    lang='pt'
)

class ExameDataExtractor:
    """Extrator otimizado de dados de exame"""

    def __init__(self):
        # Padrões regex otimizados para capturar apenas o valor específico
        self.patterns = {
            'paciente': [
                r'Paciente:\s*([^,\n\r]+?)(?:\s+Idade:|$)',
                r'Paciente:\s*([^\n\r]+)'
            ],
            'idade': [
                r'Idade:\s*(\d+\s*anos?)',
                r'Idade:\s*(\d+)'
            ],
            'nascimento': [
                r'Nascimento:\s*(\d{2}/\d{2}/\d{4})',
                r'Nascimento:\s*(\d{1,2}/\d{1,2}/\d{4})'
            ],
            'sexo': [
                r'Sexo:\s*([MF])',
                r'Sexo:\s*([MmFf])'
            ],
            'convenio': [
                r'Conv[eê]nio:\s*([^\s]+)(?:\s|$)',
                r'Conv[eê]nio:\s*([A-Z]+)'
            ],
            'prontuario': [
                r'Prontu[aá]rio:\s*(\d+)',
                r'Prontu[aá]rio:\s*([^\s]+)'
            ],
            'atendimento': [
                r'Atendimento:\s*(\d+)',
                r'Atendimento:\s*([^\s]+)'
            ],
            'numero_exame': [
                r'N[uú]mero do Exame:\s*([A-Z0-9\-/]+)',
                r'N[uú]mero do Exame:\s*([^\s]+)'
            ],
            'medico': [
                r'Dr\(a\):\s*([^,\n\r]+?)(?:\s+Categoria:|$)',
                r'Dr\(a\):\s*([A-Z\s]+?)(?:\s+[A-Z]+:|$)'
            ],
            'categoria': [
                r'Categoria:\s*([^,\n\r]+?)(?:\s+RELAT|$)',
                r'Categoria:\s*([A-Za-z\s]+?)(?:\s+[A-Z]+|$)'
            ],
            'data_entrada': [
                r'Data Entrada:\s*(\d{2}/\d{2}/\d{4})',
                r'Data Entrada:\s*(\d{1,2}/\d{1,2}/\d{4})'
            ],
            'data_liberacao': [
                r'Data Libera[cç][aã]o:\s*(\d{2}/\d{2}/\d{4})',
                r'Data Libera[cç][aã]o:\s*(\d{1,2}/\d{1,2}/\d{4})'
            ]
        }

    def extract_data_from_texts(self, textos: list) -> Dict:
        """Extrai dados estruturados da lista de textos"""
        dados = {campo: None for campo in self.patterns.keys()}

        # Criar texto unificado para regex
        texto_completo = ' '.join([self._clean_text(t) for t in textos])

        # Buscar por padrões
        for campo, patterns in self.patterns.items():
            if dados[campo]:  # Se já encontrou, pular
                continue

            for pattern in patterns:
                match = re.search(pattern, texto_completo, re.IGNORECASE | re.MULTILINE)
                if match:
                    dados[campo] = match.group(1).strip()
                    break

        # Tratamento especial para campos que podem estar na mesma linha
        for texto in textos:
            texto_limpo = self._clean_text(texto).strip()

            # Paciente na mesma linha
            if texto_limpo.startswith('Paciente:') and not dados['paciente']:
                nome = texto_limpo.replace('Paciente:', '').strip()
                if nome and not any(x in nome.lower() for x in ['idade', 'sexo', 'nascimento']):
                    dados['paciente'] = nome

            # Convênio que pode estar colado
            elif 'Convênio:' in texto_limpo or 'Convenio:' in texto_limpo:
                if not dados['convenio']:
                    parts = texto_limpo.split(':')
                    if len(parts) > 1:
                        convenio = parts[1].strip()
                        if convenio:
                            dados['convenio'] = convenio

        return dados

    def _clean_text(self, texto: str) -> str:
        """Limpa e corrige encoding do texto"""
        if not texto:
            return ""

        # Correções de encoding comuns
        corrections = {
            'Ã§': 'ç', 'Ã£': 'ã', 'Ã¡': 'á', 'Ã©': 'é',
            'Ã­': 'í', 'Ã³': 'ó', 'Ãº': 'ú', 'Ã¢': 'â',
            'Ãª': 'ê', 'Ã"': 'Ó', 'Ãš': 'Ú', 'Ã‰': 'É'
        }

        for old, new in corrections.items():
            texto = texto.replace(old, new)

        return texto

    def is_complete(self, dados: Dict) -> bool:
        """Verifica se todos os dados essenciais foram encontrados"""
        essential_fields = ['paciente', 'numero_exame', 'convenio']
        return all(dados.get(field) for field in essential_fields)


def extract_texts_from_ocr(resultado_ocr) -> list:
    """Extrai textos do resultado OCR de forma otimizada"""
    if not resultado_ocr or len(resultado_ocr) == 0:
        return []

    resultado_primeira_pagina = resultado_ocr[0]

    # Tentar acessar textos de diferentes formas
    if hasattr(resultado_primeira_pagina, 'rec_texts'):
        return resultado_primeira_pagina.rec_texts
    elif isinstance(resultado_primeira_pagina, dict) and 'rec_texts' in resultado_primeira_pagina:
        return resultado_primeira_pagina['rec_texts']
    elif isinstance(resultado_primeira_pagina, list):
        # Formato [[[x1,y1],[x2,y2],[x3,y3],[x4,y4]], (text, confidence)]
        return [item[1][0] for item in resultado_primeira_pagina if len(item) > 1]

    return []


def process_single_file_optimized(file_path: str, filename: str) -> Optional[Dict]:
    """Processa um arquivo individual de forma otimizada"""
    print(f"Processando: {filename}")
    extractor = ExameDataExtractor()

    try:
        if filename.lower().endswith('.pdf'):
            # Converter apenas a primeira página (dados geralmente estão lá)
            pages = convert_from_path(file_path, first_page=1, last_page=1)
            if not pages:
                print(f"  Erro: Nenhuma página encontrada em {filename}")
                return None

            page = pages[0]

        else:  # Imagem
            page = Image.open(file_path)

        # Converter para numpy array
        img_np = np.array(page)

        # Aplicar OCR
        resultado_ocr = ocr.predict(img_np)

        # Extrair textos
        textos = extract_texts_from_ocr(resultado_ocr)

        if not textos:
            print(f"  Aviso: Nenhum texto extraído de {filename}")
            return None

        # Extrair dados estruturados
        dados_exame = extractor.extract_data_from_texts(textos)

        # Verificar se encontrou dados essenciais
        if not extractor.is_complete(dados_exame):
            print(f"  Aviso: Dados incompletos em {filename}")
            # Tentar processar segunda página se for PDF
            if filename.lower().endswith('.pdf'):
                print(f"  Tentando segunda página...")
                try:
                    pages = convert_from_path(file_path, first_page=2, last_page=2)
                    if pages:
                        img_np = np.array(pages[0])
                        resultado_ocr = ocr.predict(img_np)
                        textos_p2 = extract_texts_from_ocr(resultado_ocr)
                        textos.extend(textos_p2)
                        dados_exame = extractor.extract_data_from_texts(textos)
                except:
                    pass

        # Resultado final
        resultado = {
            "arquivo_origem": filename,
            "data_processamento": datetime.now().isoformat(),
            "dados_exame": dados_exame,
            "status": "completo" if extractor.is_complete(dados_exame) else "incompleto"
        }

        # Salvar apenas o resultado essencial
        base_name = os.path.splitext(filename)[0]
        json_path = os.path.join(RESULTADOS_DIR, f"{base_name}_dados.json")

        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(resultado, f, ensure_ascii=False, indent=2)

        print(f"  ✅ Dados extraídos e salvos: {json_path}")

        # Log dos dados encontrados
        dados_encontrados = [k for k, v in dados_exame.items() if v]
        print(f"  📋 Campos encontrados: {', '.join(dados_encontrados)}")

        return resultado

    except Exception as e:
        print(f"  ❌ Erro ao processar {filename}: {str(e)}")
        return None


def process_all_files_optimized():
    """Processa todos os arquivos da pasta de forma otimizada"""
    print("🚀 Iniciando processamento OCR otimizado")
    print(f"📁 Diretório: {EXAMES_DIR}")
    print("-" * 50)

    if not os.path.exists(EXAMES_DIR):
        print(f"❌ Diretório {EXAMES_DIR} não encontrado!")
        return []

    resultados = []
    arquivos_processados = 0

    # Buscar arquivos válidos
    valid_extensions = ('.pdf', '.png', '.jpg', '.jpeg', '.tiff', '.bmp')
    arquivos = [f for f in os.listdir(EXAMES_DIR)
                if f.lower().endswith(valid_extensions) and
                os.path.isfile(os.path.join(EXAMES_DIR, f))]

    print(f"📊 Encontrados {len(arquivos)} arquivos para processar")
    print("-" * 50)

    for filename in arquivos:
        file_path = os.path.join(EXAMES_DIR, filename)
        resultado = process_single_file_optimized(file_path, filename)

        if resultado:
            resultados.append(resultado)
            arquivos_processados += 1

        print("-" * 30)

    # Resumo final
    print(f"🎉 Processamento concluído!")
    print(f"📊 Arquivos processados: {arquivos_processados}/{len(arquivos)}")
    print(f"💾 Resultados salvos em: {RESULTADOS_DIR}")

    # Estatísticas de qualidade
    completos = sum(1 for r in resultados if r['status'] == 'completo')
    incompletos = len(resultados) - completos

    print(f"✅ Extrações completas: {completos}")
    print(f"⚠️ Extrações incompletas: {incompletos}")

    return resultados


def extract_exam_data_from_file(file_path: str) -> Optional[Dict]:
    """
    Função simplificada para extrair dados de um único arquivo
    Retorna apenas os dados essenciais do exame
    """
    filename = os.path.basename(file_path)
    return process_single_file_optimized(file_path, filename)


def main():
    """Função principal"""
    return process_all_files_optimized()


if __name__ == "__main__":
    main()