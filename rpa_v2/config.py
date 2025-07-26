# Configurações do Sistema de Automação

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

# Configurações de tempo (em segundos)
TIMEOUTS = {
    'element_wait': 10,
    'login_redirect': 3,
    'page_load': 2,
    'search_result': 3,
    'manual_login_check': 2
}

PATIENT_NAME = "Yuri Rodrigues NEVES"
