class BaseModule:
    """Classe base para todos os módulos do sistema RPA."""

    def __init__(self, nome: str):
        self.nome = nome

    def run(self, *args, **kwargs):
        """Método que deve ser implementado pelos módulos concretos."""
        raise NotImplementedError("O método run deve ser implementado pelo módulo.")

