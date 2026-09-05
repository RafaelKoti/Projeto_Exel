# OpenpyXL e Pandas são as bibliotecas que vamo usar para ler dados das planilhas do Exel (aparentemente elas são as principais)

import openpyxl
import pandas as pd

'''Como criar as planilhas? - com essa biblioteca OpenpyXL, vamo usar o WorkBook para criar as nossas planilhas'''
criar_planilha = openpyxl.workbook()

'''Como vamos ler os arquivos Exel? - Vamos usar o Pandas (PD, pois está como abreviação da biblioteca'''

''' No parênteses vamos colocar o nome da planilha que o usuário quiser ler, mais pra frente da pra fazer um códifo onde a
váriavel mesmo lê qual arquivo a pessoa digitou'''

ler_planilhas = pd.read_excel()