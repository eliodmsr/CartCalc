#!/usr/bin/env python
# coding: utf-8

# In[2]:


#Importação de bibliotecas necessárias
import streamlit as st
import fdb
import numpy as np
import pandas as pd
import sqlalchemy
import random
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
from itertools import islice


# In[3]:


#Criando a engine de conexão e consulta ao banco de dados do servidor
con = fdb.connect(dsn='*******************', user='*******', password='*********')
cur = con.cursor()


# In[4]:


#Lista de funcionários retirada do sistema
selNomes = """SELECT ID_USUARIO as "Funcionário", NOME_FUNCIONARIO as "Funcionários", CASE WHEN PERMISSAO_ASSINAR = 'True' THEN 1 ELSE 0 END AS "Adicional Escrevente"
FROM FUNCIONARIOS
INNER JOIN TBL$USUARIO
ON FUNCIONARIOS.ID_FUNCIONARIO = TBL$USUARIO.ID_PROFISSIONAL
WHERE IDENTIFICADOR_REGIME_TRAB = 'CLT'
AND DATA_SAIDA IS NULL
AND BLOQUEADO = 'False'
ORDER BY NOME_FUNCIONARIO"""
    
cur.execute(selNomes)
dfNomes = pd.read_sql_query(selNomes, con)
dfNomes.set_index('Funcionário', inplace=True)

#função para separa listas
def split_list(a_list):
    half = len(a_list)//2
    return a_list[:half], a_list[half:]


# In[12]:


########Layout do app

########Configurações e header
st.set_page_config(layout="wide", page_icon=":books:", page_title="Calculadora de salários")
st.image('Cart.png', use_column_width='auto')
st.title(':scroll: Calculadora de salários do 1º Ofício de Notas de Montes Claros/MG')

########Colunas para definição da data
st.header('Defina a data')

colAno, colMes = st.columns(2)

with colAno:
    opcaoAno = st.number_input('Insira o ano', min_value=2020, max_value=2030, value=2021, step=1)

meses = {1:'Janeiro', 2:'Fevereiro', 3:'Março', 4:'Abril', 5:'Maio', 6:'Junho', 7:'Julho',
         8:'Agosto', 9:'Setembro', 10:'Outubro', 11:'Novembro', 12:'Dezembro'}

def format_meses(opcaoMes):
    return meses[opcaoMes]

with colMes:
    opcaoMes = st.selectbox("Selecione o mês", options=list(meses.keys()), format_func=format_meses)
    
########Coluna para definição da margem de participação
st.header('Defina a margem da produtividade de cada setor')
colAut, colEsc, colInv, colProc, colDec = st.columns(5)

with colAut:
    margAut = st.number_input('Margem setor de Autenticações', min_value=0.1, max_value=5.0, value=3.5, step=0.1)
           
with colEsc: 
    margEsc = st.number_input('Margem setor de Escrituras', min_value=0.1, max_value=5.0, value=1.5, step=0.1)

with colInv:
    margInv = st.number_input('Margem setor de Inventários', min_value=0.1, max_value=5.0, value=1.5, step=0.1)

with colProc: 
    margProc = st.number_input('Margem setor de Procurações', min_value=0.1, max_value=5.0, value=4.5, step=0.1)

with colDec:
    margDec = st.number_input('Margem setor de Declarações', min_value=0.1, max_value=5.0, value=4.0, step=0.1)

st.header('Defina o salário base e o adicional de escrevente')

colSalb, coladEsc = st.columns(2)
########Valor do salário base
with colSalb:
    salarioBase = st.number_input('Insira o valor do salário base', min_value=0.0, max_value=3000.0, value=1225.0, step=1.0)

########Adicional de Escrevente
with coladEsc:
    adEscrevente = st.number_input('Insira o valor do adicional de escrevente', min_value=100.0, max_value=1000.0, value=183.75, step=0.01)


#######Colunas de adicionais de adicionais de salário
st.header('Defina os adicionais individuais')

funcionarios = dfNomes['Funcionários']


#Produtividade do administrativo
st.subheader("Produtividade do administrativo")

gratAdmin = st.multiselect("Selecione os funcionários", funcionarios, key='admin')

selproAdmin = """select sum(ITEM_RECIBO.taxa_1)
              from ITEM_RECIBO
              inner join RECIBO on (RECIBO.ID_RECIBO = ITEM_RECIBO.ID_RECIBO)
              where (RECIBO.ARQUIVADO = 'True')
              and extract(YEAR from recibo.data_arquivado) = EXTRACT(year from CURRENT_DATE)
              and extract(MONTH from recibo.data_arquivado) = EXTRACT(month from CURRENT_DATE)"""

cur.execute(selproAdmin)
proDF = pd.read_sql_query(selproAdmin, con)
if not gratAdmin:
    preAdmin = (((proDF['SUM'].iloc[0] - (proDF['SUM'].iloc[0] * 0.275)) * 0.0067) / 1)
else:
    preAdmin = (((proDF['SUM'].iloc[0] - (proDF['SUM'].iloc[0] * 0.275)) * 0.0067) / len(gratAdmin))
                
                
valorAdmin = st.number_input('Defina o valor', min_value=0.0, max_value=2000.0, value=preAdmin, step=1.0, key='valor produtividade administrativo')

colsAdmin = st.columns(5)
valuesAdmin = []

if len(gratAdmin) > 5:
    gratAdmin1, gratAdmin2 = split_list(gratAdmin)
    for i, j in enumerate(gratAdmin1):
        valuesAdmin.append(colsAdmin[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso admin'))
    for i, j in enumerate(gratAdmin2):
        valuesAdmin.append(colsAdmin[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso admin'))
else:
    for i, j in enumerate(gratAdmin):
        valuesAdmin.append(colsAdmin[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso admin'))


dictAdmin = dict(zip(gratAdmin, valuesAdmin))


#Gratificação Escrituras
st.subheader("Gratificação Escrituras")

gratEsc = st.multiselect("Selecione os funcionários", funcionarios, key='grat esc')

valorEsc = st.number_input('Defina o valor', min_value=0.0, max_value=2000.0, value=122.50, step=1.0, key='valor esc')

colsEscrituras = st.columns(5)
valuesEsc = []


if len(gratEsc) > 5:
    gratEsc1, gratEsc2 = split_list(gratEsc)
    colsEscrituras = st.columns(5)
    for i, j in enumerate(gratEsc1):
        valuesEsc.append(colsEscrituras[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso esc0'))
    for i, j in enumerate(gratEsc2):
        valuesEsc.append(colsEscrituras[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso esc1'))
else:
    for i, j in enumerate(gratEsc):
        valuesEsc.append(colsEscrituras[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso esc2'))

        
dictEsc = dict(zip(gratEsc, valuesEsc))


#Gratificação Procurações
st.subheader("Gratificação Procurações")

gratProc = st.multiselect("Selecione os funcionários", funcionarios, key='grat proc')

valorProc = st.number_input('Defina o valor', min_value=0.0, max_value=2000.0, value=122.50, step=1.0, key='valor proc')

colsProc = st.columns(5)
valuesProc = []


if len(gratProc) > 5:
    gratProc1, gratProc2 = split_list(gratProc)
    colsProc = st.columns(5)
    for i, j in enumerate(gratProc1):
        valuesProc.append(colsProc[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso proc0'))
    for i, j in enumerate(gratProc2):
        valuesProc.append(colsProc[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso proc1'))
else:
    for i, j in enumerate(gratProc):
        valuesProc.append(colsProc[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso proc2'))

        
dictProc = dict(zip(gratProc, valuesProc))

#Função gerência
st.subheader("Função de gerência")

gratGer = st.multiselect("Selecione os funcionários", funcionarios, key='grat ger')

valorGer = st.number_input('Defina o valor', min_value=0.0, max_value=2000.0, value=1041.25, step=1.0, key='valor ger')

colsGer = st.columns(5)
valuesGer = []


if len(gratGer) > 5:
    gratGer1, gratGer2 = split_list(gratGer)
    colsGer = st.columns(5)
    for i, j in enumerate(gratGer1):
        valuesGer.append(colsGer[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso ger0'))
    for i, j in enumerate(gratGer2):
        valuesGer.append(colsGer[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso ger1'))
else:
    for i, j in enumerate(gratGer):
        valuesGer.append(colsGer[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso ger2'))

        
dictGer = dict(zip(gratGer, valuesGer))

#Gratificação comitê
st.subheader("Gratificação comitê")

gratCom = st.multiselect("Selecione os funcionários", funcionarios, key='grat com')

valorCom = st.number_input('Defina o valor', min_value=0.0, max_value=2000.0, value=612.50, step=1.0, key='valor com')

colsCom = st.columns(5)
valuesCom = []


if len(gratCom) > 5:
    gratCom1, gratCom2 = split_list(gratCom)
    colsCom = st.columns(5)
    for i, j in enumerate(gratCom1):
        valuesCom.append(colsCom[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso com0'))
    for i, j in enumerate(gratCom2):
        valuesCom.append(colsCom[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso com1'))
else:
    for i, j in enumerate(gratCom):
        valuesCom.append(colsCom[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso com2'))

        
dictCom = dict(zip(gratCom, valuesCom))


#Tarefa pessoal
st.subheader("Tarefa pessoal")

gratTap = st.multiselect("Selecione os funcionários", funcionarios, key='grat tap')

valorTap = st.number_input('Defina o valor', min_value=0.0, max_value=2000.0, value=102.08, step=1.0, key='valor tap')

colsTap = st.columns(5)
valuesTap = []


if len(gratTap) > 5:
    gratTap1, gratTap2 = split_list(gratTap)
    colsTap = st.columns(5)
    for i, j in enumerate(gratTap1):
        valuesTap.append(colsTap[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso tap0'))
    for i, j in enumerate(gratTap2):
        valuesTap.append(colsTap[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso tap1'))
else:
    for i, j in enumerate(gratTap):
        valuesTap.append(colsTap[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso tap2'))

        
dictTap = dict(zip(gratTap, valuesTap))


#Salário família
st.subheader("Salário família")

gratFam = st.multiselect("Selecione os funcionários", funcionarios, key='grat fam')

valorFam = st.number_input('Defina o valor', min_value=0.0, max_value=2000.0, value=51.27, step=1.0, key='valor fam')

colsFam = st.columns(5)
valuesFam = []


if len(gratFam) > 5:
    gratFam1, gratFam2 = split_list(gratFam)
    colsFam = st.columns(5)
    for i, j in enumerate(gratFam1):
        valuesFam.append(colsFam[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso fam0'))
    for i, j in enumerate(gratFam2):
        valuesFam.append(colsFam[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso fam1'))
else:
    for i, j in enumerate(gratFam):
        valuesFam.append(colsFam[i].number_input(f"""Insira o peso de {j}""", min_value=1.0, max_value=10.0, value=1.0, step=0.1, key='peso fam2'))

        
dictFam = dict(zip(gratFam, valuesFam))



#######Colunas fator atitudinal

st.header('Defina o fator atitudinal')

labels = dfNomes['Funcionários'].tolist()

length_to_split = [5, 5, 5, 5, 5, 5]

inputt = iter(labels)
output = [list(islice(inputt, elem))
          for elem in length_to_split]

cols0 = st.columns(len(output[0]))

atitude0 = []

for i in output[0]:
    atitude0.append(cols0[output[0].index(i)].number_input(f"""{i}""", min_value=1.0, max_value=5.0, value=1.0, step=1.0))

cols1 = st.columns(len(output[1]))

atitude1 = []

for i in output[1]:
    atitude1.append(cols0[output[1].index(i)].number_input(f"""{i}""", min_value=1.0, max_value=5.0, value=1.0, step=1.0))
    
cols2 = st.columns(len(output[2]))

atitude2 = []

for i in output[2]:
    atitude2.append(cols0[output[2].index(i)].number_input(f"""{i}""", min_value=1.0, max_value=5.0, value=1.0, step=1.0))

cols3 = st.columns(len(output[3]))

atitude3 = []

for i in output[3]:
    atitude3.append(cols0[output[3].index(i)].number_input(f"""{i}""", min_value=1.0, max_value=5.0, value=1.0, step=1.0))

cols4 = st.columns(len(output[4]))

atitude4 = []

for i in output[4]:
    atitude4.append(cols0[output[4].index(i)].number_input(f"""{i}""", min_value=1.0, max_value=5.0, value=1.0, step=1.0))

atitude = atitude0 + atitude1 + atitude2 + atitude3 + atitude4

atitudinal = []

for x in atitude:
    atitudinal.append(x / 100.00)

    
#######Colunas horas extras
st.header('Defina as horas extras')    
    
colsh0 = st.columns(len(output[0]))

horas0 = []

for i in output[0]:
    horas0.append(colsh0[output[0].index(i)].number_input(f"""{i}""", min_value=0.0, max_value=1000.0, value=0.0, step=1.0))

colsh1 = st.columns(len(output[1]))

horas1 = []

for i in output[1]:
    horas1.append(colsh0[output[1].index(i)].number_input(f"""{i}""", min_value=0.0, max_value=1000.0, value=0.0, step=1.0))
    
colsh2 = st.columns(len(output[2]))

horas2 = []

for i in output[2]:
    horas2.append(colsh0[output[2].index(i)].number_input(f"""{i}""", min_value=0.0, max_value=1000.0, value=0.0, step=1.0))

colsh3 = st.columns(len(output[3]))

horas3 = []

for i in output[3]:
    horas3.append(colsh0[output[3].index(i)].number_input(f"""{i}""", min_value=0.0, max_value=1000.0, value=0.0, step=1.0))

colsh4 = st.columns(len(output[4]))

horas4 = []

for i in output[4]:
    horas4.append(colsh0[output[4].index(i)].number_input(f"""{i}""", min_value=0.0, max_value=1000.0, value=0.0, step=1.0))

horasExtras = horas0 + horas1 + horas2 + horas3 + horas4


# In[13]:


#######Botão para gerar a planilha  
if st.button('Gerar salários'):
    #Queries para formar o data frame final
    
    #Dataframe do setor de autenticações e outros
    selAut = f"""WITH LISTA_FINAL AS (SELECT RECEPCAO_TITULO.ID_USUARIO, CAST(sum(ITEM_MOVIMENTO_CAIXA.QTDE_ATO) AS FLOAT) as "TOTAL_ATOS", sum(ITEM_MOVIMENTO_CAIXA.TAXA_1) as "VALORES"
    from MOVIMENTO_CAIXA
    INNER JOIN ITEM_MOVIMENTO_CAIXA ON ITEM_MOVIMENTO_CAIXA.ID_MOV_CAIXA = MOVIMENTO_CAIXA.ID_MOV_CAIXA
    INNER JOIN RECEPCAO_TITULO ON MOVIMENTO_CAIXA.ID_RECEPCAO = RECEPCAO_TITULO.ID_RECEPCAO
    LEFT JOIN TRANSACAO_ATO ON (TRANSACAO_ATO.ID_TRANSACAO = ITEM_MOVIMENTO_CAIXA.ID_TRANSACAO) 
    LEFT JOIN GRUPO_ATO ON (GRUPO_ATO.CODIGO = TRANSACAO_ATO.ESAT_GRAT_CODIGO)
    where GRUPO_ATO.CODIGO IN (3, 8, 46, 33, 62, 55, 5)
    and extract(YEAR from MOVIMENTO_CAIXA.DATA) = {opcaoAno}
    and extract(MONTH from MOVIMENTO_CAIXA.DATA) = {opcaoMes}
    GROUP BY RECEPCAO_TITULO.ID_USUARIO)
    SELECT ID_USUARIO as "Funcionário", TOTAL_ATOS "Quantidade de atos", 
    ((SELECT SUM(VALORES) FROM LISTA_FINAL) - ((SELECT SUM(VALORES) FROM LISTA_FINAL) * 0.275)) * 
    (TOTAL_ATOS / (SELECT SUM(TOTAL_ATOS) FROM LISTA_FINAL)) * {margAut} / 100 as "Produtividade"
    FROM LISTA_FINAL"""
    cur.execute(selAut)
    dfAut = pd.read_sql_query(selAut, con)
    dfAut.set_index('Funcionário', inplace=True)
    dfAut = dfAut.replace(to_replace= dfAut['Produtividade'].max(), value= dfAut['Produtividade'].max() * 1.1)
    
    #Dataframe para serviços de apostilamento
    selApo = f"""WITH LISTA_FINAL AS (SELECT RECEPCAO_TITULO.ID_USUARIO, CAST(sum(ITEM_MOVIMENTO_CAIXA.QTDE_ATO) AS FLOAT) as "TOTAL_ATOS", sum(ITEM_MOVIMENTO_CAIXA.TAXA_1) as "VALORES"
    from MOVIMENTO_CAIXA
    INNER JOIN ITEM_MOVIMENTO_CAIXA ON ITEM_MOVIMENTO_CAIXA.ID_MOV_CAIXA = MOVIMENTO_CAIXA.ID_MOV_CAIXA
    INNER JOIN RECEPCAO_TITULO ON MOVIMENTO_CAIXA.ID_RECEPCAO = RECEPCAO_TITULO.ID_RECEPCAO
    LEFT JOIN TRANSACAO_ATO ON (TRANSACAO_ATO.ID_TRANSACAO = ITEM_MOVIMENTO_CAIXA.ID_TRANSACAO) 
    LEFT JOIN GRUPO_ATO ON (GRUPO_ATO.CODIGO = TRANSACAO_ATO.ESAT_GRAT_CODIGO)
    where GRUPO_ATO.CODIGO IN (63)
    and extract(YEAR from MOVIMENTO_CAIXA.DATA) = {opcaoAno}
    and extract(MONTH from MOVIMENTO_CAIXA.DATA) = {opcaoMes}
    GROUP BY RECEPCAO_TITULO.ID_USUARIO)
    SELECT ID_USUARIO as "Funcionário", TOTAL_ATOS "Quantidade de atos", 
    ((SELECT SUM(VALORES) FROM LISTA_FINAL) - ((SELECT SUM(VALORES) FROM LISTA_FINAL) * 0.275)) * 
    (TOTAL_ATOS / (SELECT SUM(TOTAL_ATOS) FROM LISTA_FINAL)) * {margAut} / 100 as "Produtividade"
    FROM LISTA_FINAL"""
    cur.execute(selApo)
    dfApo = pd.read_sql_query(selApo, con)
    dfApo.set_index('Funcionário', inplace=True)
    
    
    #Dataframe Setor de escrituras (conferência e lavraturas)
    
    selConfEsc = f"""WITH ANOTACOES_IGNORADAS AS (    SELECT RECEPCAO_ANOTACOES.ID_RECEPCAO, COUNT(*) as "QUANTIDADE"
    FROM RECEPCAO_ANOTACOES 
    INNER JOIN LIVROS_NOTAS ON RECEPCAO_ANOTACOES.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO 
    WHERE EXTRACT(MONTH FROM DATA_REG) = {opcaoMes}
    AND EXTRACT(YEAR FROM DATA_REG) = {opcaoAno}
    AND ANOTACAO LIKE 'Conferid%'
    AND RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    AND LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    GROUP BY RECEPCAO_ANOTACOES.ID_RECEPCAO
    HAVING COUNT(*) > 1),
    
    LISTA_FINAL AS (SELECT RECEPCAO_ANOTACOES.ID_USUARIO, CAST(count(distinct MOVIMENTO_CAIXA.ID_RECEPCAO) AS FLOAT) as "TOTAL_ATOS", 
    sum(ITEM_MOVIMENTO_CAIXA.TAXA_1) as "VALORES"
    FROM MOVIMENTO_CAIXA 
    INNER JOIN ITEM_MOVIMENTO_CAIXA ON MOVIMENTO_CAIXA.ID_MOV_CAIXA = ITEM_MOVIMENTO_CAIXA.ID_MOV_CAIXA
    INNER JOIN RECEPCAO_ANOTACOES ON MOVIMENTO_CAIXA.ID_RECEPCAO = RECEPCAO_ANOTACOES.ID_RECEPCAO 
    INNER JOIN LIVROS_NOTAS ON MOVIMENTO_CAIXA.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO
    WHERE situacao_pgto = 'Liquidado'
    and RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    and livros_notas.ato IN (270, 94, 203, 272, 343, 216, 311, 229, 84, 130, 280, 244, 306, 352, 79, 207, 372, 88, 355, 350, 
    228, 208, 201, 100, 290, 188, 382, 131, 288, 293, 129, 113, 312, 303, 371, 313, 199, 369, 341, 89, 248, 361, 195, 365, 185, 
    297, 287, 305, 187, 86, 389, 101, 227, 356, 359, 308, 291, 395, 120, 357, 127)
    and ANOTACAO like 'Conferid%'
    and LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    and extract(YEAR from LIVROS_NOTAS.DATA_REG) = {opcaoAno}
    and extract(MONTH from LIVROS_NOTAS.DATA_REG) = {opcaoMes}
    and MOVIMENTO_CAIXA.ID_RECEPCAO NOT IN (IIF ((SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS) IS NOT NULL, 
    (SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS), (0)))
    group by RECEPCAO_ANOTACOES.ID_USUARIO)
    
    SELECT ID_USUARIO as "Funcionário", TOTAL_ATOS "Quantidade de atos", ((SELECT SUM(VALORES) FROM LISTA_FINAL) - 
    ((SELECT SUM(VALORES) FROM LISTA_FINAL) * 0.275)) * (TOTAL_ATOS / (SELECT SUM(TOTAL_ATOS) FROM LISTA_FINAL)) * 
    {margEsc} / 100 as "Produtividade"
    FROM LISTA_FINAL"""
    selLavEsc = f""" WITH ANOTACOES_IGNORADAS AS (SELECT RECEPCAO_ANOTACOES.ID_RECEPCAO, COUNT(*) as "QUANTIDADE"
    FROM RECEPCAO_ANOTACOES
    INNER JOIN LIVROS_NOTAS ON RECEPCAO_ANOTACOES.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO
    WHERE EXTRACT(MONTH FROM DATA_REG) = {opcaoMes}
    AND EXTRACT(YEAR FROM DATA_REG) = {opcaoAno}
    AND ANOTACAO LIKE 'Lavrad%'
    and RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    and LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    GROUP BY RECEPCAO_ANOTACOES.ID_RECEPCAO
    HAVING COUNT(*) > 1), 
    LISTA_FINAL AS (SELECT RECEPCAO_ANOTACOES.ID_USUARIO, CAST(count(distinct MOVIMENTO_CAIXA.ID_RECEPCAO) AS FLOAT) as "TOTAL_ATOS", 
    sum(ITEM_MOVIMENTO_CAIXA.TAXA_1) as "VALORES"
    FROM MOVIMENTO_CAIXA 
    INNER JOIN ITEM_MOVIMENTO_CAIXA ON MOVIMENTO_CAIXA.ID_MOV_CAIXA = ITEM_MOVIMENTO_CAIXA.ID_MOV_CAIXA
    INNER JOIN RECEPCAO_ANOTACOES ON MOVIMENTO_CAIXA.ID_RECEPCAO = RECEPCAO_ANOTACOES.ID_RECEPCAO
    INNER JOIN LIVROS_NOTAS ON MOVIMENTO_CAIXA.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO
    WHERE situacao_pgto = 'Liquidado'
    and RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    and livros_notas.ato IN (270, 94, 203, 272, 343, 216, 311, 229, 84, 130, 280, 244, 306, 352, 79, 207, 372, 88, 355, 350, 
    228, 208, 201, 100, 290, 188, 382, 131, 288, 293, 129, 113, 312, 303, 371, 313, 199, 369, 341, 89, 248, 361, 195, 365, 185, 
    297, 287, 305, 187, 86, 389, 101, 227, 356, 359, 308, 291, 395, 120, 357, 127)
    and ANOTACAO like 'Lavrad%'
    and LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    and extract(YEAR from LIVROS_NOTAS.DATA_REG) = {opcaoAno}
    and extract(MONTH from LIVROS_NOTAS.DATA_REG) = {opcaoMes}
    and MOVIMENTO_CAIXA.ID_RECEPCAO NOT IN (IIF ((SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS) IS NOT NULL, 
    (SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS), (0)))
    group by RECEPCAO_ANOTACOES.ID_USUARIO)
    SELECT ID_USUARIO as "Funcionário", TOTAL_ATOS "Quantidade de atos", 
    ((SELECT SUM(VALORES) FROM LISTA_FINAL) - ((SELECT SUM(VALORES) FROM LISTA_FINAL) * 0.275)) * 
    (TOTAL_ATOS / (SELECT SUM(TOTAL_ATOS) FROM LISTA_FINAL)) * {margEsc} / 100 as "Produtividade"
    FROM LISTA_FINAL"""
    cur.execute(selLavEsc)
    cur.execute(selConfEsc)
    dfLavEsc = pd.read_sql_query(selLavEsc, con)
    dfConfEsc = pd.read_sql_query(selConfEsc, con)
    dfEsc = pd.concat([dfLavEsc, dfConfEsc]).groupby(['Funcionário']).sum()
    dfEsc = dfEsc.replace(to_replace= dfEsc['Produtividade'].max(), value= dfEsc['Produtividade'].max() * 1.1)   
    
    #Dataframe Setor de inventários (conferência e lavraturas)
    
    selConfInv = f"""WITH ANOTACOES_IGNORADAS AS (SELECT RECEPCAO_ANOTACOES.ID_RECEPCAO, COUNT(*) as "QUANTIDADE"
    FROM RECEPCAO_ANOTACOES
    INNER JOIN
    LIVROS_NOTAS ON RECEPCAO_ANOTACOES.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO
    WHERE EXTRACT(MONTH FROM DATA_REG) = {opcaoMes}
    AND EXTRACT(YEAR FROM DATA_REG) = {opcaoAno}
    AND ANOTACAO LIKE 'Conferid%'
    and RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    and LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    GROUP BY RECEPCAO_ANOTACOES.ID_RECEPCAO
    HAVING COUNT(*) > 1),
    LISTA_FINAL AS (SELECT RECEPCAO_ANOTACOES.ID_USUARIO, CAST(count(distinct MOVIMENTO_CAIXA.ID_RECEPCAO) AS FLOAT) as "TOTAL_ATOS", sum(ITEM_MOVIMENTO_CAIXA.TAXA_1) as "VALORES"
    FROM MOVIMENTO_CAIXA 
    INNER JOIN 
    ITEM_MOVIMENTO_CAIXA ON MOVIMENTO_CAIXA.ID_MOV_CAIXA = ITEM_MOVIMENTO_CAIXA.ID_MOV_CAIXA
    INNER JOIN
    RECEPCAO_ANOTACOES ON MOVIMENTO_CAIXA.ID_RECEPCAO = RECEPCAO_ANOTACOES.ID_RECEPCAO
    INNER JOIN
    LIVROS_NOTAS ON MOVIMENTO_CAIXA.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO
    WHERE situacao_pgto = 'Liquidado'
    and RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    and livros_notas.ato IN (309, 344, 321, 260, 186, 346, 370, 198, 374, 80, 323, 377, 226, 349, 304, 284, 302, 249, 146, 194, 232, 106, 173)
    and ANOTACAO like 'Conferid%'
    and LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    and extract(YEAR from LIVROS_NOTAS.DATA_REG) = {opcaoAno}
    and extract(MONTH from LIVROS_NOTAS.DATA_REG) = {opcaoMes}
    and MOVIMENTO_CAIXA.ID_RECEPCAO NOT IN (IIF ((SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS) IS NOT NULL, (SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS), (0)))
    group by RECEPCAO_ANOTACOES.ID_USUARIO)
    SELECT ID_USUARIO as "Funcionário", TOTAL_ATOS "Quantidade de atos", 
    ((SELECT SUM(VALORES) FROM LISTA_FINAL) - ((SELECT SUM(VALORES) FROM LISTA_FINAL) * 0.275)) * 
    (TOTAL_ATOS / (SELECT SUM(TOTAL_ATOS) FROM LISTA_FINAL)) * {margInv} / 100 as "Produtividade"
    FROM LISTA_FINAL"""

    selLavInv = f"""WITH ANOTACOES_IGNORADAS AS (SELECT RECEPCAO_ANOTACOES.ID_RECEPCAO, COUNT(*) as "QUANTIDADE"
    FROM RECEPCAO_ANOTACOES
    INNER JOIN
    LIVROS_NOTAS ON RECEPCAO_ANOTACOES.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO
    WHERE EXTRACT(MONTH FROM DATA_REG) = {opcaoMes}
    AND EXTRACT(YEAR FROM DATA_REG) = {opcaoAno}
    AND ANOTACAO LIKE 'Lavrad%'
    and RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    and LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    GROUP BY RECEPCAO_ANOTACOES.ID_RECEPCAO
    HAVING COUNT(*) > 1),
    LISTA_FINAL AS (SELECT RECEPCAO_ANOTACOES.ID_USUARIO, CAST(count(distinct MOVIMENTO_CAIXA.ID_RECEPCAO) AS FLOAT) as "TOTAL_ATOS", sum(ITEM_MOVIMENTO_CAIXA.TAXA_1) as "VALORES"
    FROM MOVIMENTO_CAIXA 
    INNER JOIN 
    ITEM_MOVIMENTO_CAIXA ON MOVIMENTO_CAIXA.ID_MOV_CAIXA = ITEM_MOVIMENTO_CAIXA.ID_MOV_CAIXA
    INNER JOIN
    RECEPCAO_ANOTACOES ON MOVIMENTO_CAIXA.ID_RECEPCAO = RECEPCAO_ANOTACOES.ID_RECEPCAO
    INNER JOIN
    LIVROS_NOTAS ON MOVIMENTO_CAIXA.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO
    WHERE situacao_pgto = 'Liquidado'
    and RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    and livros_notas.ato IN (309, 344, 321, 260, 186, 346, 370, 198, 374, 80, 323, 377, 226, 349, 304, 284, 302, 249, 146, 194, 232, 106, 173)
    and ANOTACAO like 'Lavrad%'
    and LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    and extract(YEAR from LIVROS_NOTAS.DATA_REG) = {opcaoAno}
    and extract(MONTH from LIVROS_NOTAS.DATA_REG) = {opcaoMes}
    and MOVIMENTO_CAIXA.ID_RECEPCAO NOT IN (IIF ((SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS) IS NOT NULL, (SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS), (0)))
    group by RECEPCAO_ANOTACOES.ID_USUARIO)
    SELECT ID_USUARIO as "Funcionário", TOTAL_ATOS "Quantidade de atos", 
    ((SELECT SUM(VALORES) FROM LISTA_FINAL) - ((SELECT SUM(VALORES) FROM LISTA_FINAL) * 0.275)) * 
    (TOTAL_ATOS / (SELECT SUM(TOTAL_ATOS) FROM LISTA_FINAL)) * {margInv} / 100 as "Produtividade"
    FROM LISTA_FINAL"""

    cur.execute(selLavInv)
    cur.execute(selConfInv)

    dfLavInv = pd.read_sql_query(selLavInv, con)
    dfConfInv = pd.read_sql_query(selConfInv, con)

    dfInv = pd.concat([dfLavInv, dfConfInv]).groupby(['Funcionário']).sum()
    dfInv = dfInv.replace(to_replace= dfInv['Produtividade'].max(), value= dfInv['Produtividade'].max() * 1.1)
    
    
    #Dataframe Setor de procurações

    selProc = f"""WITH ANOTACOES_IGNORADAS AS (SELECT RECEPCAO_ANOTACOES.ID_RECEPCAO, COUNT(*) as "QUANTIDADE"
    FROM RECEPCAO_ANOTACOES
    INNER JOIN
    LIVROS_NOTAS ON RECEPCAO_ANOTACOES.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO
    WHERE EXTRACT(MONTH FROM DATA_REG) = {opcaoMes}
    AND EXTRACT(YEAR FROM DATA_REG) = {opcaoAno}
    AND ANOTACAO LIKE 'Lavrad%'
    and RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    and LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    GROUP BY RECEPCAO_ANOTACOES.ID_RECEPCAO
    HAVING COUNT(*) > 1),
    LISTA_FINAL as (SELECT RECEPCAO_ANOTACOES.ID_USUARIO, CAST(count(distinct MOVIMENTO_CAIXA.ID_RECEPCAO) AS FLOAT) as "TOTAL_ATOS", sum(ITEM_MOVIMENTO_CAIXA.TAXA_1) as "VALORES"
    FROM MOVIMENTO_CAIXA 
    INNER JOIN 
    ITEM_MOVIMENTO_CAIXA ON MOVIMENTO_CAIXA.ID_MOV_CAIXA = ITEM_MOVIMENTO_CAIXA.ID_MOV_CAIXA
    INNER JOIN
    RECEPCAO_ANOTACOES ON MOVIMENTO_CAIXA.ID_RECEPCAO = RECEPCAO_ANOTACOES.ID_RECEPCAO
    INNER JOIN
    LIVROS_NOTAS ON MOVIMENTO_CAIXA.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO
    WHERE situacao_pgto = 'Liquidado'
    and RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    and livros_notas.ato IN (58, 7, 264, 237, 23, 316, 20, 256, 257, 250, 39, 383, 235, 6, 386, 271, 236, 385, 238, 251, 222, 259, 
    215, 217, 21, 25, 24, 390, 176, 241, 22, 384, 401)
    and ANOTACAO like 'Lavrad%'
    and LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    and extract(YEAR from LIVROS_NOTAS.DATA_REG) = {opcaoAno}
    and extract(MONTH from LIVROS_NOTAS.DATA_REG) = {opcaoMes}
    and MOVIMENTO_CAIXA.ID_RECEPCAO NOT IN (IIF ((SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS) IS NOT NULL, (SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS), (0)))
    group by RECEPCAO_ANOTACOES.ID_USUARIO)
    SELECT ID_USUARIO as "Funcionário", TOTAL_ATOS "Quantidade de atos", 
    ((SELECT SUM(VALORES) FROM LISTA_FINAL) - ((SELECT SUM(VALORES) FROM LISTA_FINAL) * 0.275)) * 
    (TOTAL_ATOS / (SELECT SUM(TOTAL_ATOS) FROM LISTA_FINAL)) * {margProc} / 100 as "Produtividade"
    FROM LISTA_FINAL"""

    cur.execute(selProc)
    dfProc = pd.read_sql_query(selProc, con)
    dfProc.set_index('Funcionário', inplace=True)
    dfProc = dfProc.replace(to_replace= dfProc['Produtividade'].max(), value= dfProc['Produtividade'].max() * 1.1)
    
    #Dataframe Setor de Declarações
    
    selDec = f"""WITH ANOTACOES_IGNORADAS AS (SELECT RECEPCAO_ANOTACOES.ID_RECEPCAO, COUNT(*) as "QUANTIDADE"
    FROM RECEPCAO_ANOTACOES
    INNER JOIN
    LIVROS_NOTAS ON RECEPCAO_ANOTACOES.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO
    WHERE EXTRACT(MONTH FROM DATA_REG) = {opcaoMes}
    AND EXTRACT(YEAR FROM DATA_REG) = {opcaoAno}
    AND ANOTACAO LIKE 'Lavrad%'
    and RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    and LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    GROUP BY RECEPCAO_ANOTACOES.ID_RECEPCAO
    HAVING COUNT(*) > 1),
    LISTA_FINAL AS (SELECT RECEPCAO_ANOTACOES.ID_USUARIO, CAST(count(distinct MOVIMENTO_CAIXA.ID_RECEPCAO) AS FLOAT) as "TOTAL_ATOS", sum(ITEM_MOVIMENTO_CAIXA.TAXA_1) as "VALORES"
    FROM MOVIMENTO_CAIXA 
    INNER JOIN 
    ITEM_MOVIMENTO_CAIXA ON MOVIMENTO_CAIXA.ID_MOV_CAIXA = ITEM_MOVIMENTO_CAIXA.ID_MOV_CAIXA
    INNER JOIN
    RECEPCAO_ANOTACOES ON MOVIMENTO_CAIXA.ID_RECEPCAO = RECEPCAO_ANOTACOES.ID_RECEPCAO
    INNER JOIN
    LIVROS_NOTAS ON MOVIMENTO_CAIXA.ID_RECEPCAO = LIVROS_NOTAS.ID_RECEPCAO
    WHERE situacao_pgto = 'Liquidado'
    and RECEPCAO_ANOTACOES.EXCLUIDO = 'False'
    and livros_notas.ato IN (242, 57, 335, 252, 398, 268, 218, 399, 108, 266, 110, 107, 267, 262, 179, 263, 275, 269, 274, 109, 
    172, 225, 230, 298, 180, 273, 147, 265, 178, 261, 245, 111)
    and ANOTACAO like 'Lavrad%'
    and LIVROS_NOTAS.ID_LIVRO_INICIO IS NOT NULL
    and extract(YEAR from LIVROS_NOTAS.DATA_REG) = {opcaoAno}
    and extract(MONTH from LIVROS_NOTAS.DATA_REG) = {opcaoMes}
    and MOVIMENTO_CAIXA.ID_RECEPCAO NOT IN (IIF ((SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS) IS NOT NULL, (SELECT LIST(ID_RECEPCAO) FROM ANOTACOES_IGNORADAS), (0)))
    group by RECEPCAO_ANOTACOES.ID_USUARIO)
    SELECT ID_USUARIO as "Funcionário", TOTAL_ATOS "Quantidade de atos", 
    ((SELECT SUM(VALORES) FROM LISTA_FINAL) - ((SELECT SUM(VALORES) FROM LISTA_FINAL) * 0.275)) * 
    (TOTAL_ATOS / (SELECT SUM(TOTAL_ATOS) FROM LISTA_FINAL)) * {margDec} / 100 as "Produtividade"
    FROM LISTA_FINAL"""

    cur.execute(selDec)
    dfDec = pd.read_sql_query(selDec, con)
    dfDec.set_index('Funcionário', inplace=True)
    dfDec = dfDec.replace(to_replace= dfDec['Produtividade'].max(), value= dfDec['Produtividade'].max() * 1.1)
    
    #######Adicionando gratificações
    
    #Produtividade do administrativo 
    grat1 = []
    
    for row in dfNomes['Funcionários']:
        if row in gratAdmin:
            grat1.append(valorAdmin)       
        else:
            grat1.append(0.00)
    
    dfNomes['Produtividade administrativo'] = grat1
                
    for key, value in dictAdmin.items():
        dfNomes['Produtividade administrativo'] = np.where(dfNomes['Funcionários'] == key, dfNomes['Produtividade administrativo'] * value, dfNomes['Produtividade administrativo'])       
    
    #Gratificação Escrituras
    
    grat2 = []
    
    for row in dfNomes['Funcionários']:
        if row in gratEsc:
            grat2.append(valorEsc)       
        else:
            grat2.append(0.00)
    
    dfNomes['Gratificação Escrituras'] = grat2
                
    for key, value in dictEsc.items():
        dfNomes['Gratificação Escrituras'] = np.where(dfNomes['Funcionários'] == key, dfNomes['Gratificação Escrituras'] * value, dfNomes['Gratificação Escrituras'])
    
    #Gratificação Procurações
    
    grat3 = []
    
    for row in dfNomes['Funcionários']:
        if row in gratProc:
            grat3.append(valorProc)       
        else:
            grat3.append(0.00)
    
    dfNomes['Gratificação Procurações'] = grat3
                
    for key, value in dictProc.items():
        dfNomes['Gratificação Procurações'] = np.where(dfNomes['Funcionários'] == key, dfNomes['Gratificação Procurações'] * value, dfNomes['Gratificação Procurações'])
    
    #Função Gerência
    
    grat4 = []
    
    for row in dfNomes['Funcionários']:
        if row in gratGer:
            grat4.append(valorGer)       
        else:
            grat4.append(0.00)
    
    dfNomes['Função de Gerência'] = grat4
                
    for key, value in dictGer.items():
        dfNomes['Função de Gerência'] = np.where(dfNomes['Funcionários'] == key, dfNomes['Função de Gerência'] * value, dfNomes['Função de Gerência'])
 

    #Gratificação comitê
    
    grat5 = []
    
    for row in dfNomes['Funcionários']:
        if row in gratCom:
            grat5.append(valorCom)       
        else:
            grat5.append(0.00)
    
    dfNomes['Gratificação comitê'] = grat5
                
    for key, value in dictCom.items():
        dfNomes['Gratificação comitê'] = np.where(dfNomes['Funcionários'] == key, dfNomes['Gratificação comitê'] * value, dfNomes['Gratificação comitê'])
    
    #Tarefa pessoal
    
    grat6 = []
    
    for row in dfNomes['Funcionários']:
        if row in gratTap:
            grat6.append(valorTap)       
        else:
            grat6.append(0.00)
    
    dfNomes['Tarefa pessoal'] = grat6
                
    for key, value in dictTap.items():
        dfNomes['Tarefa pessoal'] = np.where(dfNomes['Funcionários'] == key, dfNomes['Tarefa pessoal'] * value, dfNomes['Tarefa pessoal'])
   
    #Salário família
    
    grat7 = []
    
    for row in dfNomes['Funcionários']:
        if row in gratFam:
            grat7.append(valorFam)       
        else:
            grat7.append(0.00)
    
    dfNomes['Salário Família'] = grat7
                
    for key, value in dictFam.items():
        dfNomes['Salário Família'] = np.where(dfNomes['Funcionários'] == key, dfNomes['Salário Família'] * value, dfNomes['Salário Família'])

    
    
    ######Gerando Dataframe final
    df = pd.concat([dfAut, dfApo, dfEsc, dfInv, dfProc, dfDec]).groupby(['Funcionário']).sum().sort_values(by=['Funcionário']).join(dfNomes, how='right').set_index(['Funcionários']).drop('Quantidade de atos', axis=1)
    
    df['Adicional Escrevente'] = df['Adicional Escrevente'].replace(1, adEscrevente)    
    
    df.insert(0, 'Salário base', salarioBase)
    
    df = df.fillna(0).round(2)
    
    df['Horas Extras'] = horasExtras
    
    df['Produtividade'] = df['Produtividade'] + (df['Produtividade'] * atitudinal)
    
    df['Salário'] = df[list(df.columns)].sum(axis=1)
    

    #Função para habilitar download da tabela para uma planilha do Excel
    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data

    df_xlsx = to_excel(df)
    
    #Tabela-resultado
    st.dataframe(df)
    st.download_button(label='📥 Baixar planilha',
                                    data=df_xlsx ,
                                    file_name= 'Salários 1º Ofício.xlsx')

else:
    st.write('*Preencha os dados e aperte o botão para iniciar*')

