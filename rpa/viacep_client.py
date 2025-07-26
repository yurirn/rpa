import requests
from urllib.parse import quote
BASE_URL = "https://viacep.com.br/ws"

class ViaCEPError(Exception):
    """Exceção lançada quando a consulta não retorna dados válidos."""

def buscar_endereco(uf: str, cidade: str, logradouro: str, timeout: int = 10) -> list[dict]:
    """
    Consulta o ViaCEP por UF, cidade e logradouro.
    
    Args:
        uf (str): Sigla do estado (ex.: 'PR').
        cidade (str): Nome da cidade (ex.: 'Londrina').
        logradouro (str): Rua, avenida, etc. (ex.: 'Rua César de Oliveira Bertin').
        timeout (int, opcional): Tempo máximo (s) de espera pela resposta. Padrão: 10 s.

    Returns:
        list[dict]: Lista de endereços (pode vir mais de um CEP quando a rua é longa).

    Raises:
        ViaCEPError: Se não encontrar nada ou resposta for inválida.
        requests.HTTPError / requests.Timeout: Erros de rede normais.
    """
    url = f"{BASE_URL}/{uf}/{quote(cidade)}/{quote(logradouro)}/json/"
    response = requests.get(url, timeout=timeout)
    response.raise_for_status() 

    data = response.json()

    if isinstance(data, dict) and data.get("erro"):
        raise ViaCEPError("Endereço não encontrado.")
    if not isinstance(data, list):
        raise ViaCEPError("Resposta inesperada da API ViaCEP.")

    return data