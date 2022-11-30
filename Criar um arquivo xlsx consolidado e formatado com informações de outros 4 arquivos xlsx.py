#Criar um arquivo xlsx consolidado e formatado com informações de outros 4 arquivos xlsx

import pandas as opcoesDoPanda
import os

#Caminho onde estão os arquivos
caminhoArquivo = r"C:\Users\Windows\Desktop\Python Projetos\Consolidando Arquivos\Excel"

#Variável onde estão todos os arquivos
listaArquivos = os.listdir(caminhoArquivo)

#Listando todos os arquivos + o caminho
listaCaminhoEArquivoExcel = [caminhoArquivo + '\\' + arquivo for arquivo in listaArquivos if arquivo[-4:] == 'xlsx' ]

#Criando um DataFrame
dadosArquivo = opcoesDoPanda.DataFrame()

#Copiando todos os dados dos arquivos em DadosArquivo
for arquivo in listaCaminhoEArquivoExcel:
    dados = opcoesDoPanda.read_excel(arquivo)
    dadosArquivo = dadosArquivo.append(dados) #.append sera substituido por .concat no futuro. Ignorar a msg de erro.

#Criando uma nova planilha e passando os dados dos arquivos
dadosArquivo.to_excel(r"C:\Users\Windows\Desktop\Python Projetos\Consolidando Arquivos\ArquivoConsolidado.xlsx")

#######################################################################################################################

#Formatando a planilha Dados Consolidados:

from openpyxl import load_workbook
from openpyxl import workbook

caminhoArquivoDadosSistema = "C:\\Users\\Windows\\Desktop\\Python Projetos\\Consolidando Arquivos\ArquivoConsolidado.xlsx"
planilhaDadosSistema = load_workbook(filename=caminhoArquivoDadosSistema)

#Seleciona a Sheet1
sheetPlanilhaDadosSistema = planilhaDadosSistema['Sheet1']

#Deletando a coluna A
sheetPlanilhaDadosSistema.delete_cols(1)

#Renomeando a Sheet1 para Dados Consolidados
sheetPlanilhaDadosSistema.title = 'Dados Consolidados'

#Ajustando a largura das colunas A e B
sheetPlanilhaDadosSistema.column_dimensions['A'].width = 35
sheetPlanilhaDadosSistema.column_dimensions['B'].width = 40

#Para salvar as alterações em segundo plano
planilhaDadosSistema.save(filename=caminhoArquivoDadosSistema)

#Para abrir o arquivo em primeiro plano
os.startfile(caminhoArquivoDadosSistema)