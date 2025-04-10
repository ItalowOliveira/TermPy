from datetime import date

data_atual = date.today()

data_em_texto = "0{}/0{}/{}".format(data_atual.day, data_atual.month, data_atual.year)

print(data_em_texto)