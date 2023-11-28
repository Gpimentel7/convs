import pandas as pd

def converter_sped_para_xlsx(arquivo_txt, arquivo_xlsx):
    # Ler o arquivo SPED em formato de texto
    with open(arquivo_txt, 'r') as file:
        linhas = file.readlines()

    # Criar um dicionário para armazenar os dados de cada registro
    dados = {}

    # Processar cada linha do arquivo SPED
    for linha in linhas:
        campos = linha.strip().split('|')  # Considerando o exemplo em que os campos são separados por '|'
        registro = campos[1]  # O segundo campo contém o código do registro
        if registro not in dados:
            dados[registro] = []
        dados[registro].append(campos)

    # Criar um escritor do Excel
    writer = pd.ExcelWriter(arquivo_xlsx, engine='xlsxwriter')

    # Salvar cada registro em uma planilha separada no arquivo Excel
    for registro, valores in dados.items():
        df = pd.DataFrame(valores, columns=None)
        df.to_excel(writer, sheet_name=registro, index=False, header=False)

    # Fechar o escritor para salvar o arquivo Excel
    writer.close()

# Exemplo de uso
arquivo_txt = r'C:\Users\gustavo.pimentel\Documents\sped\SpedEPC-44115676000104-Original-out2023.txt'
arquivo_xlsx = r'C:\Users\gustavo.pimentel\Documents\sped\SpedEPC-44115676000104-Original-out2023.xlsx'
converter_sped_para_xlsx(arquivo_txt, arquivo_xlsx)
