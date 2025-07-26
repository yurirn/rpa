# Configurações do Sistema de Automação

# Configurações da Interface
GUI_TITLE = "Sistema de Automação - Logs"
GUI_GEOMETRY = "800x600"
GUI_FONT_TITLE = ("Arial", 14, "bold")
GUI_FONT_STATUS = ("Arial", 10, "bold")
GUI_FONT_LOG = ("Consolas", 9)

# Configurações do Chrome
CHROME_OPTIONS = [
    "--start-maximized",
    "--disable-blink-features=AutomationControlled"
]

# Seletores e identificadores do sistema web
SELECTORS = {
    'username_field': "username",
    'password_field': "password",
    'login_button': "//button[@type='submit'] | //input[@type='submit']",
    'create_exam_button': "//a[contains(text(), 'Criar novo exame')]",
    'modal_create_button': "//div[@id='modalCriarNovoExame']//a[contains(text(), 'Criar novo exame')]",
    'patient_search_field': "pacienteSearch",
    'consult_button': "consultarPaciente",
    'create_patient': "criarPaciente",
    'patient_selected': "usarPacienteSelecionado",

    'data_nascimento': "dataNascmentoDoPacienteFluxo",
    'idade': "idadeDoPacienteFluxo",
    'celular': "telefone",
    'nome_nascimento': "nomeNascimento",

    'cep': "cep"

}

# URLs e padrões para verificação de login
LOGIN_SUCCESS_PATTERNS = [
    '/moduloexame/index'
]

LOGIN_FAIL_PATTERNS = [
    '/login/auth'
]

# Configurações de tempo (em segundos)
TIMEOUTS = {
    'element_wait': 10,
    'login_redirect': 3,
    'page_load': 2,
    'search_result': 3,
    'manual_login_check': 2
}

# Nome do paciente (configurável)
PATIENT_NAME = "Yuri Rodrigues NEVES"  

# Cores para logs
LOG_COLORS = {
    'success': 'green',
    'error': 'red',
    'info': 'blue',
    'warning': 'orange'
}