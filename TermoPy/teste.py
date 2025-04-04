import pandas as pd

caminho_arquivo = "equipamentos.xlsx"

CodEquipamento = 'AEC055'

op1 = 2

if op1 == 1:

    df = pd.read_excel(caminho_arquivo, sheet_name="Computadores", header=1)
    selected_columns = ['NOME', 'MODELO', 'N/S', 'PREÇO']
    linha = df.loc[df['NOME'] == CodEquipamento, selected_columns]

elif op1 == 2:

    df = pd.read_excel(caminho_arquivo, sheet_name="Equipamentos", header=1)
    selected_columns = ['NOME', 'DESCRIÇÃO', 'MODELO', 'N/S', 'PREÇO']
    linha = df.loc[df['NOME'] == CodEquipamento, selected_columns]

elif op1 == 3:

    df = pd.read_excel(caminho_arquivo, sheet_name="Celulares", header=1)
    selected_columns = ['NOME', 'DESCRIÇÃO', 'MODELO', 'N/S', 'PREÇO']
    linha = df.loc[df['NOME'] == CodEquipamento, selected_columns]



print("Visualização geral dos dados:")
print(linha)





#df = pd.read_excel(caminho_arquivo, sheet_name="Computadores")
#selected_columns = ['NOME', 'MODELO', 'N/S', 'PRECO']
#
#df = pd.read_excel(caminho_arquivo, sheet_name="Celulares")
#selected_columns = ['NOME', 'DESCRIÇÃO', 'MODELO', 'N/S', 'PREÇO']