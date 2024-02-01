import pandas as pd
import reader
import re
import os  # biblioteca usada para utilização do metodo de listagem do nome dos arquivos

caminhopasta = "C:\\Users\\ramon\\OneDrive\\Documentos\\GitHub\\analisadorPlanilhas\\planilhas\\"  # necessário implementar container para que o caminho da pasta se mantenha o mesmo idependente da máquina que o código roda
#caminhopasta = "C:\\Users\\ramon\\Documents\\GitHub\\analisadorPlanilhas\\planilhas\\" # pcTrabalho

nomeArquivos = os.listdir(caminhopasta) # Lista todos os arquivos na pasta e armazena em uma array

planilha = pd.read_excel('Compilado.xlsx', sheet_name='database')


numeroLinhasColuna = len(reader.colecaoPlanilhas[0]['Aparelho gela?']) # o metodo len conta o numero de linhas de colecaoPlanilhas

for i in range(1,numeroLinhasColuna):
    planilha.at[i, "Aparelho gela?"] = reader.colecaoPlanilhas[0]['Aparelho gela?'][i]
   
   
planilha.to_excel('Compilado.xlsx', sheet_name='database', index=False)



print('fim do programa')




# linha = 0
# coluna = "A"
# carro = planilha.iat[0,0]
# print(carro)