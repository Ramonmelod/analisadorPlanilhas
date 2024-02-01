import pandas as pd
from openpyxl import load_workbook  # load_workbook permite abrir e carregar um arquivo Excel existente para manipulação.
import re
import os  # biblioteca usada para utilização do metodo de listagem do nome dos arquivos

caminhopasta = "C:\\Users\\ramon\\OneDrive\\Documentos\\GitHub\\analisadorPlanilhas\\planilhas\\"  # necessário implementar container para que o caminho da pasta se mantenha o mesmo idependente da máquina que o código roda
#caminhopasta = "C:\\Users\\ramon\\Documents\\GitHub\\analisadorPlanilhas\\planilhas\\" # pcTrabalho
colecaoPlanilhas =[]

nomeArquivos = os.listdir(caminhopasta) # Lista todos os arquivos na pasta e armazena em uma array

planilha = pd.read_excel('Compilado.xlsx', sheet_name='database')



for i in range(1,15):
    planilha.at[i, "Aparelho gela?"] = "não"
   
   
planilha.to_excel('Compilado.xlsx', sheet_name='database', index=False)





print('fim do programa')




# linha = 0
# coluna = "A"
# carro = planilha.iat[0,0]
# print(carro)