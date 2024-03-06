def sobrescrever_dicionario(dicionario_original, dicionario_substituto):
    for chave, valor in dicionario_substituto.items():
        if chave in dicionario_original:
            # Se a chave já existe no dicionário original, substitua o valor
            dicionario_original[chave] = valor
        else:
            # Se a chave não existe, adicione-a ao dicionário original
            dicionario_original[chave] = valor


# Exemplo de uso:
dicionario_original = {'a': 1, 'b': 2, 'c': 3}
dicionario_substituto = {'b': 5, 'c': 7, 'd': 9}

sobrescrever_dicionario(dicionario_original, dicionario_substituto)

print(dicionario_original)
