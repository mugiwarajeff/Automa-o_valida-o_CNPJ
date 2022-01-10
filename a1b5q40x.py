import requests
import sys
import pandas as pd
from time import sleep
import re
import json
import openpyxl

planilha_retorno_final = openpyxl.load_workbook("cnpjs.xlsx")
planilha_retorno_pagina_trabalho = planilha_retorno_final['Plan1']
print(planilha_retorno_final)
def tratar_dados(elemento_lista):   #usando regex para tratar os dados
    return re.sub(r'[^0-9]', '', elemento_lista)

def cnpjval(file):
    conteudo = pd.read_excel(file, header=None, dtype=str, squeeze=True)
    return conteudo


df = cnpjval('cnpjs.xlsx')


cnpjs = [tratar_dados(cnpj) for cnpj in df]
lista = []


'''
with open('output.txt', 'w') as f:
        for count, item in enumerate(cnpjs, start=1):
            cnpj = item
            try:
                print("Acessando CNPJ {} ({} de {})...".format(cnpj, count, len(cnpjs)))
                url = f'https://www.receitaws.com.br/v1/cnpj/{cnpj}'
                response = requests.request('GET', url)
                response_lista = json.loads(response.text)
                print(response_lista["porte"])
                lista.append(response.text)
                f.write(response_lista["porte"])
            except Exception as e:
                f.write(f"O CNPJ: {cnpj} não é valido!... {e}")
                print(f"O CNPJ: {cnpj} não é valido!... {e}")
            sleep(21)
'''



