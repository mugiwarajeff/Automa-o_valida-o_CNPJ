from openpyxl.workbook import workbook
import requests
import sys
import pandas as pd
from time import sleep
import re
import json
import openpyxl


def tratar_dados(elemento_lista):   #usando regex para tratar os dados
    return re.sub(r'[^0-9]', '', elemento_lista)

def cnpjval(file):
    conteudo = pd.read_excel(file, header=None, dtype=str, squeeze=True)
    return conteudo

def transporte_planilha(lista):
    planilha_retorno_final = openpyxl.load_workbook("cnpjs.xlsx")
    planilha_retorno_pagina_trabalho = planilha_retorno_final['Plan1']
    for numero, item in enumerate(lista):
        planilha_retorno_pagina_trabalho.cell(row=(numero+1), column=2, value=item)
    
    planilha_retorno_final.save("cnpjs.xlsx")

df = cnpjval('cnpjs.xlsx')

cnpjs = [tratar_dados(cnpj) for cnpj in df]
lista_logs = []

for count, item in enumerate(cnpjs, start=1):
    cnpj = item
    try:
        print("Acessando CNPJ {} ({} de {})...".format(cnpj, count, len(cnpjs)))
        url = f'https://www.receitaws.com.br/v1/cnpj/{cnpj}'
        response = requests.request('GET', url)
        response_lista = json.loads(response.text)
        print(response_lista["porte"])
        lista_logs.append(response_lista["porte"])
    except Exception as e:
        print(f"O CNPJ: {cnpj} não é valido!... {e}")
        lista_logs.append(f"O CNPJ: {cnpj} não é valido!... {e}")
    sleep(21)

try:
    transporte_planilha(lista_logs)
except Exception as e: 
    print("limpe a segunda coluna da planilha")
