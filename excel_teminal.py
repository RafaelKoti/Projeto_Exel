# OpenpyXL e Pandas são as bibliotecas que vamo usar para ler dados das planilhas do Exel (aparentemente elas são as principais)
import openpyxl
import pandas as pd

# Código criar tabela
def criar_planilha():
    criar_planilha = openpyxl.Workbook()
    aba = criar_planilha.active  # pega a aba ativa
    aba['A1'] = "Teste"
    criar_planilha.save("meu_arquivo.xlsx")

# Código para ler planilhas usando o Pandas
#ler_planilhas = pd.read_excel()

#A = Coluna
#1 = linha