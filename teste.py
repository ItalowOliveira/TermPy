<<<<<<< HEAD
import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title("Tabela de Dados - Agrofértil")
root.geometry("600x300")

frame = ttk.Frame(root)
frame.pack(fill='both', expand=True)

colunas = ("ID", "Produto", "Quantidade", "Preço")

tabela = ttk.Treeview(frame, columns=colunas, show='headings')

for coluna in colunas:
    tabela.heading(coluna, text=coluna)
    tabela.column(coluna, anchor="center")

dados = [
    (1, "Fertilizante NPK", 50, "R$ 120,00"),
    (2, "Semente de Milho", 100, "R$ 200,00"),
    (3, "Calcário", 30, "R$ 75,00")
]

for item in dados:
    tabela.insert('', 'end', values=item)

# Scrollbar vertical
scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tabela.yview)
tabela.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")

# Adicionando a tabela ao frame
tabela.pack(fill='both', expand=True)

# Loop da aplicação
root.mainloop()
=======
from datetime import date

data_atual = date.today()

data_em_texto = "0{}/0{}/{}".format(data_atual.day, data_atual.month, data_atual.year)

print(data_em_texto)
>>>>>>> 4630ef156becbf1a1130612e88d4ce5b86eb6226
