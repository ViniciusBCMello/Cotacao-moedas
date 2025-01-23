import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import requests
import pandas as pd
import numpy as np
from datetime import datetime


#obtendo o nome das moedas que são mostradas no combobox_moeda através de request para awesomeapi
requesicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requesicao.json() #transformando do formato json
lista_moedas = list(dicionario_moedas.keys()) #transformando o dicionário em uma lista e obtendo as chaves dele


def pegar_cotacao():
    moeda = combobox_moeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requesicao_moeda = requests.get(link)
    cotacao = requesicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    label_resultadocotacao['text'] = f'A cotação da moeda {moeda} no dia {data_cotacao} foi de: R${valor_moeda}'

def selecionar_excel():
    caminho_arquivo = askopenfilename(title='Selecione o Arquivo de Moeda')
    var_caminhoarquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivoselecionado['text'] = f'Arquivo Selecionado: {caminho_arquivo}'

def atualizar_cotacoes():
    try:    #garantir que o arquivo enviado seja excel
        #ler o dataframe de moedas
        df = pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:,0] #ler a primeira coluna do DataFrame

        #data de inicio e fim das cotações
        data_inicial = calendario_datainicial.get()
        data_final = calendario_datafinal.get()

        #separação da data inicial para link da requisição
        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]

        #separação da data final para link da requisição
        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        #separar e repetir o processo para cada uma das moedas contidas no DataFrame
        for moeda in moedas:
            link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}/?start_date={ano_inicial}{mes_inicial}{dia_inicial}&end_date={ano_final}{mes_final}{dia_final}'
            requesicao_moeda = requests.get(link)
            cotacoes = requesicao_moeda.json()
            #cada cotação entre as datas selecionadas
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                #valor do fechamento da moeda
                bid = float(cotacao['bid']) 
                #tratamento de data extraída da API
                data = datetime.timestamp(timestamp)
                data = datetime.strftime('%d/%m/%Y')
                #tratando a forma de transferência do arquivo para o excel
                if data not in df:
                    df[data] = np.nan #preenchimento das colunas vazias como nan para depois atualizar
                #salvar a informação onde a primeira linha corresponde a moeda e a coluna corresponde a data
                df.loc[df.iloc[:,0] == moeda ,data] = bid 
        #atualizando o arquivo no excel
        df.to_excel("Moedas.xlsx")
        label_arquivoatualizado['text'] = 'Arquivo Atualizado com Sucesso'  
    except: #caso o arquivo náo seja excel, envia mensagem de erro
        label_arquivoatualizado['text'] = 'Selecione um Arquivo Excel no formato correto'


#Criação da janela
janela = tk.Tk()
janela.title('Ferramenta de Cotação de Moedas')

## COTAÇÃO DE UMA MOEDA ##

## 1L significa primeira linha, 1C significa primeira coluna. ##
#construção LABEL da 1L da janela
label_cotacaomoeda = tk.Label(text='Cotação de 1 moeda específica', borderwidth=2, relief='solid')
label_cotacaomoeda.grid(row=0, column=0, padx=10, pady=10, sticky='nsew', columnspan=3)

#construção LABEL da 2L da janela
label_selecionarmoeda = tk.Label(text='Selecione a moeda que deseja consultar:',anchor='e')
label_selecionarmoeda.grid(row=1, column=0, padx=10, pady=10, sticky='NSEW', columnspan=2)

#construção BOX da 2L 2C da janela
combobox_moeda = ttk.Combobox(values=lista_moedas)
combobox_moeda.grid(row=1,column=2, sticky="NSEW", padx=10, pady=10)

#construção LABEL da 3L da janela
label_selecionardia = tk.Label(text='Selecione o dia que deseja pegar a cotação:',anchor='e')
label_selecionardia.grid(row=2, column=0, padx=10, pady=10, sticky='NSEW', columnspan=2)

#construção DATE da 3L 3C da janela
calendario_moeda = DateEntry(year=2025, locale='pt_br')
calendario_moeda.grid(row=2, column=2, sticky='nsew', padx=10, pady=10)

#construção LABEL da 4L da janela
label_resultadocotacao = tk.Label(text='')
label_resultadocotacao.grid(row=3, column=0, padx=10, pady=10, sticky='NSEW', columnspan=2)

#construção BUTTON - PEGAR COTAÇÕES da 4L 3C da janela
botao_pegarcotacao = tk.Button(text='Pegar cotação', command= pegar_cotacao)
botao_pegarcotacao.grid(row=3, column=2, sticky='nsew', padx=10, pady=10)

## COTAÇÃO DE VÁRIAS MOEDAS ##

#construção LABEL da 5L da janela
label_cotacaovariasmoedas = tk.Label(text='Cotação de Múltiplas Moedas', borderwidth=2, relief='solid')
label_cotacaovariasmoedas.grid(row=4, column=0, padx=10, pady=10, sticky='NSEW', columnspan=3)

#construção LABEL da 6L da janela
label_selecionarexcel = tk.Label(text='Selecione um arquivo em Excel com as Moedas na Coluna A:')
label_selecionarexcel.grid(row=5, column=0, sticky="nsew", padx=10, pady=10, columnspan=2)

var_caminhoarquivo = tk.StringVar()

#construção BUTTON - SELECIONAR ARQUIVO da 6L 3C da janela
botao_selecionararquivo = tk.Button(text='Clique aqui para selecionar', command= selecionar_excel)
botao_selecionararquivo.grid(row=5, column=2, sticky='nsew', padx=10, pady=10)

#construção LABEL da 7L da janela
label_arquivoselecionado = tk.Label(text='Nenhum Arquivo selecionado',anchor='e')
label_arquivoselecionado.grid(row=6, column=0, sticky='nsew', padx=10, pady=10, columnspan=3)

#construção LABEL da 8L da janela
#construção LABEL da 9L da janela
label_datainicial = tk.Label(text='Data Inicial:',anchor='w')
label_datafinal = tk.Label(text='Data Final:',anchor='w')
label_datainicial.grid(row=7, column=0, sticky='nsew', padx=10, pady=10)
label_datafinal.grid(row=8, column=0, sticky='nsew', padx=10, pady=10)

#construção DATA da 8L 2C da janela
#construção DATA da 9L 2C da janela
calendario_datainicial = DateEntry(year=2025, locale='pt_br')
calendario_datafinal = DateEntry(year=2025, locale='pt_br')
calendario_datainicial.grid(row=7, column=1, sticky='nsew', padx=10, pady=10)
calendario_datafinal.grid(row=8, column=1, sticky='nsew', padx=10, pady=10)

#construção BUTTON - ATUALIZAR COTAÇÕES da 10L da janela
botao_atualizarcotacoes = tk.Button(text='Atualizar Cotações', command= atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=9, column=0, sticky='nsew', padx=10, pady=10)

#construção LABEL da 10L 2C da janela
label_arquivoatualizado = tk.Label(text='')
label_arquivoatualizado.grid(row=9, column=1, sticky='nsew', padx=10, pady=10, columnspan=2)

#construção BUTTON - FECHAR da 11L da janela
botao_fechar = tk.Button(text='Fechar', command= janela.quit)
botao_fechar.grid(row=10, column=3, sticky='nsew', padx=10, pady=10)

janela.mainloop()