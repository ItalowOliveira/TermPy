import pandas as pd
from docx import Document
import tkinter as tk

def carregar_contador():
        try:
            with open('contador.txt', 'r') as f:
                return int(f.read())
        except FileNotFoundError:
            return 0  

def salvar_contador(contador):
        with open('contador.txt', 'w') as f:
            f.write(str(contador))

def incrementar_contador():
        contador = carregar_contador()
        contador += 1
        salvar_contador(contador)
        return contador


def RetornaDados():
    CPF = Edit2.get()
    Nome = Edit1.get()
    CodEquipamento = ''
    NumeroSequencial = incrementar_contador()

    doc = Document("Termo.docx")

    for paragrafo in doc.paragraphs:
        if 'NumeroTermo' in paragrafo.text:
         paragrafo.text = paragrafo.text.replace('NumeroTermo', str(NumeroSequencial))

    for paragrafo in doc.paragraphs:
        if 'NomeColaborador' in paragrafo.text:
         paragrafo.text = paragrafo.text.replace('NomeColaborador', Nome)

    for paragrafo in doc.paragraphs:
        if 'CPFColaborador' in paragrafo.text:
         paragrafo.text = paragrafo.text.replace('CPFColaborador', CPF)
    
    caminho_arquivo = "equipamentos.xlsx"
    op1 = variavel_selecao.get()

    if op1 == 'Computador':
        CodEquipamento = 'AFPC'+Edit3.get()
        df = pd.read_excel(caminho_arquivo, sheet_name="Computadores", header=1)
        selected_columns = ['NOME', 'MODELO', 'N/S', 'PREÇO']
        linha = df.loc[df['NOME'] == CodEquipamento, selected_columns]

        if not linha.empty:
            nome = linha.iloc[0]['NOME']
            modelo = linha.iloc[0]['MODELO']
            numero_serie = linha.iloc[0]['N/S']
            preco = linha.iloc[0]['PREÇO']
            print (nome)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                     cell.text = cell.text.replace('ID_ROW1', str(nome))
                    


    elif op1 == 'Equipamento':
        CodEquipamento = 'AFEQPTO'+Edit3.get() 
        df = pd.read_excel(caminho_arquivo, sheet_name="Equipamentos", header=1)
        selected_columns = ['NOME', 'DESCRIÇÃO', 'MODELO', 'N/S', 'PREÇO']
        linha = df.loc[df['NOME'] == CodEquipamento, selected_columns]

    elif op1 == 'Celulares':
        CodEquipamento = 'AFCEL'+Edit3.get()
        df = pd.read_excel(caminho_arquivo, sheet_name="Celulares", header=1)
        selected_columns = ['NOME', 'DESCRIÇÃO', 'MODELO', 'N/S', 'PREÇO']
        linha = df.loc[df['NOME'] == CodEquipamento, selected_columns]
    else:
            print('Não escolheu')

    doc.save('Termo.Editado.docx')
    print("Visualização geral dos dados:")
    print(linha)
    return {CPF, Nome, NumeroSequencial}


















root = tk.Tk()
root.title("Gerador de Termo")
root.geometry("400x300") 

Lbl1 = tk.Label(root, text='Nome do Colaborador')
Lbl1.pack()

Edit1 = tk.Entry(root)
Edit1.pack()

Lbl2 = tk.Label(root, text='CPF:')
Lbl2.pack()

Edit2 = tk.Entry(root)
Edit2.pack()


Lbl4= tk.Label(root)

Lbl3 = tk.Label(root, text='Tipo de Equipamento:')
Lbl3.pack()

variavel_selecao = tk.StringVar()
variavel_selecao.set("Escolha uma opção")

opcoes = ["Computador", "Equipamento", "Celulares"]

menu_select = tk.OptionMenu(root, variavel_selecao, *opcoes)
menu_select.pack()

Lbl4 = tk.Label(root, text='Cod do Item')
Lbl4.pack()


Edit3 = tk.Entry(root)
Edit3.pack()

SubmitBtn = tk.Button(root,text="Gerar Termo", command=RetornaDados)
SubmitBtn .pack()

root.mainloop()










