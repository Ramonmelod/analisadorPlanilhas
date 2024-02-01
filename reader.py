import pandas as pd
from openpyxl import load_workbook  # load_workbook permite abrir e carregar um arquivo Excel existente para manipulação.
import re
import os  # biblioteca usada para utilização do metodo de listagem do nome dos arquivos

caminhopasta = "C:\\Users\\ramon\\OneDrive\\Documentos\\GitHub\\analisadorPlanilhas\\planilhas\\"  # necessário implementar container para que o caminho da pasta se mantenha o mesmo idependente da máquina que o código roda
#caminhopasta = "C:\\Users\\ramon\\Desktop\\levantamento\\planilhas\\" # pcTrabalho
colecaoPlanilhas =[]
colecaoNumeroAbas = []

nomeArquivos = os.listdir(caminhopasta) # Lista todos os arquivos na pasta e armazena em uma array



for i in range(0, 45):
   planilhaSetor = pd.ExcelFile(caminhopasta + nomeArquivos[i])
   numeroAbasPlanilha = len(planilhaSetor.sheet_names) # recebe o numero de abas da planilha i
   colecaoNumeroAbas.append(numeroAbasPlanilha)        # guarda o numero de abas da planilha i na posição i do array colecaoNumeroabas 
   #print(i, " = ----"," .O arquivo: - ",nomeArquivos[i]," possui: ---", colecaoNumeroAbas[i]," abas")

   for j in range(0,colecaoNumeroAbas[i]):
      colecaoPlanilhas.append(planilhaSetor.parse(sheet_name=j))  # como o metodo ExcelFile lê toda a planilha, se faz necessário que se aplique este for para ler todas as abas
   planilhaSetor.close()                                          # fecha o arquivo apenas quando o for com j termina


print("---------------fim do programa reader-------------")




# for k in range(0,46):                                       # imprime todas as linhas de todas as planilhas com colulas Aparelho gela?

#    print(colecaoPlanilhas[i]['Aparelho gela?'])
#    print("--------fim da planilha")

   
#print(colecaoPlanilhas[2]['Aparelho gela?'])


#    print(colecaoPlanilhas[0]['Aparelho gela?'])
# print("--------fim da planilha")
# print(colecaoPlanilhas[1]['Aparelho gela?'])
# print("--------fim da planilha")
# print(colecaoPlanilhas[2]['Aparelho gela?'])
# print("--------fim da planilha")
# print(colecaoPlanilhas[3]['Aparelho gela?'])
# print("--------fim da planilha")
# print(colecaoPlanilhas[4]['Aparelho gela?'])
