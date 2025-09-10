def formatar_cartao_17_digitos(cartao):
    """Formata o número do cartão para ter 17 dígitos, adicionando 0 antes se necessário"""
    cartao_limpo = str(cartao).strip()
    
    # Remover espaços e caracteres especiais, manter apenas números e letras
    cartao_sem_espacos = ''.join(cartao_limpo.split())
    
    if len(cartao_sem_espacos) == 16:
        cartao_formatado = "0" + cartao_sem_espacos
        print(f"📋 Cartão formatado: {cartao_sem_espacos} → {cartao_formatado}")
        return cartao_formatado
    elif len(cartao_sem_espacos) == 17:
        print(f"📋 Cartão já tem 17 dígitos: {cartao_sem_espacos}")
        return cartao_sem_espacos
    else:
        print(f"⚠️ Cartão com tamanho inesperado ({len(cartao_sem_espacos)} dígitos): {cartao_sem_espacos}")
        return cartao_sem_espacos

# Testar com os dados da planilha
cartoes_teste = [
    "005000000249273G",
    "005000000472390E", 
    "0005000000249273G"
]

print("=== TESTE DE FORMATAÇÃO DE CARTÃO ===")
for cartao in cartoes_teste:
    print(f"\nCartão original: '{cartao}' (len: {len(cartao)})")
    resultado = formatar_cartao_17_digitos(cartao)
    print(f"Resultado: '{resultado}' (len: {len(resultado)})")
    
    # Mostrar como ficaria no JavaScript
    javascript_code = f'''$("#codigo").val("{resultado}").trigger("input").trigger("change").trigger("blur");'''
    print(f"JavaScript: {javascript_code}")
    print("-" * 50) 