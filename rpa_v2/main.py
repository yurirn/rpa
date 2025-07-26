"""
Sistema RPA - Clínica
Ponto de entrada principal
"""

import sys
import os

# Adicionar src ao path para poder importar os módulos
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

# Importar a janela principal
from src.ui.main_window import MainWindow

def main():
    """Função principal que inicia a aplicação"""
    print("Iniciando Sistema RPA...")

    # Criar e executar a janela principal
    app = MainWindow()
    app.run()

    print("Sistema finalizado.")

if __name__ == "__main__":
    main()