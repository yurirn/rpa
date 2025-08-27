# 📋 Instruções para Gerar Executável (.exe) do Sistema RPA

## 🚀 Método 1: Script Automatizado (Recomendado)

### Passo 1: Executar o Script de Build
```bash
python build_exe.py
```

Este script irá:
- ✅ Instalar o PyInstaller automaticamente
- ✅ Criar arquivo de configuração personalizado
- ✅ Gerar o executável
- ✅ Criar script de instalação

### Passo 2: Instalar o Sistema
Após a geração, execute:
```bash
instalar_sistema.bat
```

## 🔧 Método 2: Manual com PyInstaller

### Passo 1: Instalar PyInstaller
```bash
pip install pyinstaller
```

### Passo 2: Gerar Executável
```bash
pyinstaller --onefile --windowed --add-data "config.json;." --add-data "modules.json;." --add-data "src;src" --add-data "xml;xml" main.py
```

## 📁 Estrutura do Executável

O executável será criado na pasta `dist/` com:
- `Sistema_RPA.exe` - Executável principal
- Todas as dependências incluídas
- Arquivos de configuração e módulos

## ⚠️ Requisitos do Sistema

- **Windows 10/11** (64-bit)
- **Python 3.8+** (apenas para gerar o .exe)
- **Memória RAM**: Mínimo 4GB, Recomendado 8GB+
- **Espaço em disco**: ~200MB para instalação

## 🎯 Características do Executável

- ✅ **Portátil**: Não requer instalação do Python
- ✅ **Standalone**: Todas as dependências incluídas
- ✅ **GUI**: Interface gráfica sem console
- ✅ **Otimizado**: Compilado para melhor performance

## 🔍 Solução de Problemas

### Erro: "Falha ao carregar módulos"
- Verifique se `modules.json` está na mesma pasta do .exe
- Verifique se a pasta `src/` está presente

### Erro: "Configuração não encontrada"
- Verifique se `config.json` está na mesma pasta do .exe

### Erro: "Dependências não encontradas"
- Execute o script `build_exe.py` novamente
- Verifique se todas as bibliotecas estão no `requirements.txt`

## 📦 Distribuição

Para distribuir o sistema:
1. Copie a pasta `dist/` completa
2. Ou execute `instalar_sistema.bat` no computador de destino
3. O sistema funcionará em qualquer Windows sem Python instalado

## 🆘 Suporte

Se encontrar problemas:
1. Verifique se todos os arquivos estão presentes
2. Execute como administrador se necessário
3. Verifique se o antivírus não está bloqueando o executável

---
**Desenvolvido para Sistema RPA - Clínica**
*Versão: 1.0.0* 