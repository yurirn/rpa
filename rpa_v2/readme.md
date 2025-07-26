# Sistema de Automação RPA - Clínica

Sistema modular de automação de processos para clínicas médicas, desenvolvido em Python com Selenium.

## 📋 Índice

- [Visão Geral](#visão-geral)
- [Arquitetura](#arquitetura)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Uso](#uso)
- [Desenvolvimento de Módulos](#desenvolvimento-de-módulos)
- [Estrutura do Projeto](#estrutura-do-projeto)

## 🎯 Visão Geral

Este sistema permite automatizar diversos processos repetitivos em sistemas de clínicas, incluindo:

- ✅ Criação de exames anatomopatológicos
- ✅ Preparação de lotes Unimed (leitura de Excel e atualização de status)
- ✅ Envio de lotes para Unimed
- 🔄 Outros módulos podem ser facilmente adicionados

### Características Principais

- **Modular**: Cada processo é um módulo independente
- **Interface Gráfica**: Seleção fácil de módulos e visualização de logs
- **Configurável**: Parâmetros ajustáveis para cada módulo
- **Extensível**: Fácil adicionar novos módulos
- **Robusto**: Tratamento de erros e logging detalhado

## 🏗️ Arquitetura

O sistema segue uma arquitetura modular em camadas:

```
┌─────────────────┐
│   Interface UI  │ ← Seleção de módulos e parâmetros
├─────────────────┤
│ Module Registry │ ← Gerenciamento de módulos
├─────────────────┤
│    Modules      │ ← Implementação das automações
├─────────────────┤
│      Core       │ ← Classes base e utilitários
├─────────────────┤
│     Browser     │ ← Selenium WebDriver
└─────────────────┘
```

## 🚀 Instalação

### Pré-requisitos

- Python 3.8 ou superior
- Google Chrome instalado
- Git (opcional)

### Passos de Instalação

1. **Clone o repositório** (ou extraia o arquivo ZIP):
```bash
git clone https://github.com/seu-usuario/clinic-automation-rpa.git
cd clinic-automation-rpa
```

2. **Crie um ambiente virtual**:
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

3. **Instale as dependências**:
```bash
pip install -r requirements.txt
```

4. **Crie o arquivo de configuração**:
```bash
cp .env.example .env
# Edite o arquivo .env com suas credenciais
```

## ⚙️ Configuração

### 1. Configuração do Sistema (.env)

Crie um arquivo `.env` na raiz do projeto:

```env
# URL do sistema
SYSTEM_URL=https://sistema-clinica.com.br

# Credenciais (opcional - pode fazer login manual)
LOGIN_USER=seu_usuario
LOGIN_PASS=sua_senha

# Outras configurações
DEBUG=False
```

### 2. Configuração da Aplicação (config/app_config.json)

O arquivo é criado automaticamente na primeira execução, mas pode ser editado:

```json
{
  "browser": {
    "headless": false,
    "timeout": 30
  },
  "logging": {
    "level": "INFO",
    "file": "logs/automation.log"
  }
}
```

### 3. Configuração de Módulos

Cada módulo tem seu próprio `config.json` em sua pasta:

```json
{
  "exam_types": {
    "175": "AN - Anátomo Patológico",
    "180": "CI - Citologia"
  },
  "default_timeout": 10
}
```

## 📖 Uso

### Executando o Sistema

1. **Inicie a aplicação**:
```bash
python main.py
```

2. **Na interface gráfica**:
   - Selecione o módulo desejado
   - Preencha os parâmetros necessários
   - Clique em "Executar"
   - Acompanhe o progresso e logs

### Módulos Disponíveis

#### 1. Criação de Exames
- **Descrição**: Automatiza a criação de exames anatomopatológicos
- **Parâmetros**:
  - Nome do Paciente
  - Data de Nascimento
  - Telefone
  - Tipo de Exame
  - Convênio
  - Médico Requisitante
  - Procedência
  - Quantidade de Material

#### 2. Preparação Lote Unimed
- **Descrição**: Lê arquivo Excel e atualiza status dos exames
- **Parâmetros**:
  - Arquivo Excel de entrada
  - Status a aplicar
  - Gerar relatório

#### 3. Envio Lote Unimed
- **Descrição**: Envia lote preparado para o sistema Unimed
- **Parâmetros**:
  - Arquivo de lote
  - Validar antes de enviar

## 🔧 Desenvolvimento de Módulos

### Criando um Novo Módulo

1. **Crie uma pasta para o módulo**:
```bash
mkdir src/modules/meu_modulo
touch src/modules/meu_modulo/__init__.py
```

2. **Implemente a classe do módulo**:

```python
# src/modules/meu_modulo/meu_modulo.py
from src.modules.base_module import AutomationModule
from src.core.base_automation import BaseAutomation

class MeuModulo(AutomationModule, BaseAutomation):
    def __init__(self, name="meu_modulo", description="", logger=None, browser_manager=None):
        super().__init__(
            name=name,
            description="Descrição do meu módulo",
            logger=logger,
            browser_manager=browser_manager
        )
    
    def validate_prerequisites(self):
        # Validar pré-requisitos
        return True, "OK"
    
    def get_parameters(self):
        # Definir parâmetros necessários
        return {
            'param1': {
                'type': 'string',
                'label': 'Parâmetro 1',
                'required': True
            }
        }
    
    def execute(self, parameters):
        # Implementar a automação
        self.update_progress(50, "Processando...")
        # ... código da automação
        return True
```

3. **Exporte a classe no __init__.py**:
```python
from .meu_modulo import MeuModulo
__all__ = ['MeuModulo']
```

4. **Reinicie a aplicação** - o módulo será descoberto automaticamente!

### Métodos Úteis da BaseAutomation

```python
# Clicar em elemento
self.click_element(By.ID, "meu-botao")

# Preencher campo
self.fill_field(By.NAME, "nome", "João Silva")

# Campo editável (com âncora)
self.activate_editable_field(
    "input#campo + a.ancora",
    "input#campo",
    "valor"
)

# Aguardar e selecionar em typeahead
self.wait_and_select_from_typeahead(
    "#medico",
    "Dr. Silva",
    "Silva"
)

# Executar sequência de ações
actions = [
    {'type': 'click', 'selector': '#btn1'},
    {'type': 'wait', 'seconds': 2},
    {'type': 'fill', 'selector': '#campo1', 'value': 'texto'}
]
self.execute_action_sequence(actions)
```

## 📁 Estrutura do Projeto

```
rpa/
├── src/
│   ├── core/                    # Classes base e utilitários
│   │   ├── browser_factory.py   
│   │   └── logger.py           
│   │
│   ├── modules/                # Módulos de automação
│   │   ├── criacao_exames.py      
│   │   ├── envio_lote_unimed.py      
│   │   ├── preparacao_lote_unimed.py      
│   │   └── preparacao_lote_unimed_novo.py     
│   │
│   ├── ui/                     # Interface gráfica
│   │   └── main_window.py      # Janela principal
│   │
│   └── utils/                  # Utilitários
│       └── viacep_client.py    # Cliente ViaCEP
│
├── .env                   
├── config.json    
├── modules.json              
├── main.py                    
├── requirements.txt           
└── README.md                  
```

## 🛠️ Solução de Problemas

### Erro: "Chrome driver não encontrado"
- O sistema baixa automaticamente o ChromeDriver
- Verifique sua conexão com a internet

### Erro: "Módulo não encontrado"
- Recarregue os módulos na interface
- Verifique se o módulo está na pasta correta

### Sistema lento
- Ajuste os timeouts em `config/app_config.json`
- Considere usar modo headless para melhor performance

## 📝 Licença

Este projeto é proprietário e confidencial.

## 👥 Suporte

Para suporte e dúvidas:
- Abra uma issue no repositório
- Entre em contato com a equipe de desenvolvimento