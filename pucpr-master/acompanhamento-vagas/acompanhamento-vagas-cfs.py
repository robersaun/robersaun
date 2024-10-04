"""Acompanhamento de vagas em câmpus fora de sede

Script para acompanhar o número de vagas em turmas durante o ajuste acadêmico.
Recebe como entrada um único relatório extraído do sistema Prime
e cria um arquivo excel para acompanhamento de vagas.
Não contêm a coluna 'salas' e 'capacidade'.

Utilizado pela equipe de informações acadêmicas.

Desenvolvido por Vinicius Tozo
Última atualização: 20/07/2021
"""
# coding:utf8
import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas as pd
import pytz
import xlsxwriter

# Janela para usuário selecionar arquivo
Tk().withdraw()
acompanhamento_vagas = askopenfilename(
    filetypes=[('Arquivo excel', '.xlsx')],
    title='Selecione o relatório de carga horária por turma / disciplina')

# Dicionários para seleção de tipo de turma e escola
dicionario_tipos = {
    "U1": "Eletiva",
    "U2": "Eletiva",
    "U3": "Eletiva",
    "E": "Especial",
    "E1": "Especial",
    "E2": "Especial",
    "GCL": "Global Class",
    "GCL1": "Global Class",
    "GCL2": "Global Class",
    "GCL3": "Global Class",
    "GCL3 1": "Global Class",
    "GCL3 2": "Global Class",
    "LEAEADBA": "LEA",
    "LEAEADCV": "LEA",
    "LEAEADDI": "LEA",
    "LEAEADEH": "LEA",
    "LEAEADME": "LEA",
    "LEAEADNE": "LEA",
    "LEAEADPO": "LEA",
    "LETTCEAD": "LETTC",
    "LETTCU1": "LETTC",
    "LETTCU2": "LETTC",
    "A": "Regular",
    "B": "Regular",
    "C": "Regular",
    "D": "Regular",
    "U": "Regular",
}
dicionario_escolas = {
    115: 'Londrina',
    2501: 'Londrina',
    111: 'Londrina',
    2502: 'Londrina',
    2518: 'Londrina',
    119: 'Londrina',
    111148: 'Londrina',
    111154: 'Londrina',
    105130: 'Londrina',
    2530: 'Londrina',
    121: 'Londrina',
    124: 'Londrina',
    114: 'Londrina',
    2610: 'Maringá',
    2611: 'Maringá',
    114140: 'Maringá',
    2605: 'Maringá',
    2606: 'Maringá',
    2603: 'Maringá',
    2614: 'Maringá',
    2615: 'Maringá',
    2402: 'Toledo',
    2420: 'Toledo',
    2424: 'Toledo',
    2404: 'Toledo',
    2429: 'Toledo',
    2430: 'Toledo',
    2431: 'Toledo',
    2422: 'Toledo',
    2410: 'Toledo',
    112152: 'Toledo',
    2423: 'Toledo',
    112153: 'Toledo',
    2281: 'Belas Artes',
    2283: 'Belas Artes',
    2287: 'Politécnica',
    2289: 'Politécnica',
    2295: 'Ciências da Vida',
    2093: 'Ciências da Vida',
    2291: 'Ciências da Vida',
    1060: 'Ciências da Vida',
    2062: 'Medicina',
    2063: 'Ciências da Vida',
    2064: 'Ciências da Vida',
    2066: 'Ciências da Vida',
    2069: 'Ciências da Vida',
    2070: 'Ciências da Vida',
    2071: 'Ciências da Vida',
    2072: 'Ciências da Vida',
    2074: 'Ciências da Vida',
    2075: 'Ciências da Vida',
    2076: 'Ciências da Vida',
    2078: 'Ciências da Vida',
    2079: 'Ciências da Vida',
    2160: 'Ciências da Vida',
    2166: 'Ciências da Vida',
    2167: 'Ciências da Vida',
    2206: 'Educação e Humanidades',
    2212: 'Ciências da Vida',
    2268: 'Educação e Humanidades',
    2311: 'Ciências da Vida',
    2312: 'Educação e Humanidades',
    2360: 'Ciências da Vida',
    2369: 'Ciências da Vida',
    2121: 'Educação e Humanidades',
    2140: 'Politécnica',
    2145: 'Politécnica',
    2147: 'Politécnica',
    2158: 'Politécnica',
    2159: 'Politécnica',
    2222: 'Politécnica',
    2223: 'Politécnica',
    2241: 'Politécnica',
    2242: 'Belas Artes',
    2243: 'Politécnica',
    2244: 'Belas Artes',
    2245: 'Politécnica',
    2246: 'Politécnica',
    2247: 'Belas Artes',
    2248: 'Politécnica',
    2249: 'Educação e Humanidades',
    2250: 'Politécnica',
    2252: 'Politécnica',
    2253: 'Politécnica',
    2254: 'Belas Artes',
    2255: 'Belas Artes',
    2256: 'Belas Artes',
    2263: 'Politécnica',
    2269: 'Politécnica',
    2270: 'Politécnica',
    2271: 'Belas Artes',
    2272: 'Belas Artes',
    2303: 'Belas Artes',
    2341: 'Politécnica',
    2342: 'Belas Artes',
    2343: 'Politécnica',
    2344: 'Politécnica',
    2345: 'Politécnica',
    2346: 'Politécnica',
    2347: 'Belas Artes',
    2348: 'Politécnica',
    2349: 'Educação e Humanidades',
    2350: 'Educação e Humanidades',
    2351: 'Politécnica',
    2352: 'Politécnica',
    2353: 'Politécnica',
    2354: 'Belas Artes',
    2355: 'Politécnica',
    2356: 'Belas Artes',
    2357: 'Belas Artes',
    2359: 'Politécnica',
    2022: 'Direito',
    2023: 'Direito',
    2031: 'Belas Artes',
    2220: 'Educação e Humanidades',
    2224: 'Belas Artes',
    2225: 'Belas Artes',
    2226: 'Belas Artes',
    2227: 'Negócios',
    2231: 'Belas Artes',
    2232: 'Belas Artes',
    2236: 'Belas Artes',
    2237: 'Belas Artes',
    2238: 'Belas Artes',
    2084: 'Direito',
    2085: 'Direito',
    2028: 'Negócios',
    2029: 'Negócios',
    2035: 'Negócios',
    2036: 'Negócios',
    2128: 'Negócios',
    2129: 'Negócios',
    2229: 'Negócios',
    2230: 'Negócios',
    2239: 'Negócios',
    2278: 'Negócios',
    2279: 'Negócios',
    2280: 'Negócios',
    2201: 'Educação e Humanidades',
    2202: 'Educação e Humanidades',
    2203: 'Educação e Humanidades',
    2210: 'Educação e Humanidades',
    2214: 'Educação e Humanidades',
    2215: 'Belas Artes',
    2219: 'Educação e Humanidades',
    2221: 'Educação e Humanidades',
    2275: 'Educação e Humanidades',
    2277: 'Educação e Humanidades',
    2293: 'Negócios',
    2301: 'Educação e Humanidades',
    2302: 'Educação e Humanidades',
    2304: 'Educação e Humanidades',
    2308: 'Belas Artes',
    2314: 'Educação e Humanidades',
    2319: 'Educação e Humanidades'
}


# Função para selecionar os dados para a coluna 'Análise'
def analisar_diferenca(diferenca):
    if diferenca > 0:
        return "Vagas Disponíveis"
    elif diferenca == 0:
        return "Não há vagas"
    elif diferenca < 0:
        return "Vagas excedidas"
    else:
        return ""


# Leitura do arquivo
df_vagas = pd.read_excel(acompanhamento_vagas)

# Ordenação das colunas
df_vagas = df_vagas[
    ["Estabelecimento", "CR Curso", "Curso", "Período", "Agrupamento", "Turma", "Disciplina", "Divisão",
     "Qtde de Alunos Matri.", "Nr Vagas Cadastradas", "Professor", "Data da Semana", "Tem Corte"]]

df_vagas = df_vagas.loc[df_vagas['Data da Semana'] != " "]
df_vagas = df_vagas.loc[df_vagas['Tem Corte'] == "Sim"]

# Conversão de colunas para inteiro
df_vagas["Período"] = df_vagas["Período"].fillna(0.0).astype(int)
df_vagas["Qtde de Alunos Matri."] = df_vagas["Qtde de Alunos Matri."].replace(' ', 0.0).fillna(0.0) \
    .astype(int)
df_vagas["Nr Vagas Cadastradas"] = df_vagas["Nr Vagas Cadastradas"].replace(' ', 0.0).fillna(0.0).astype(int)
df_vagas["CR Curso"] = df_vagas["CR Curso"].replace(' ', 0.0).fillna(0.0).astype(int)

# Criação da coluna 'Escola', com os valores do dicionário conforme o código do curso
df_vagas.insert(1, 'Escola', '', True)
df_vagas['Escola'] = df_vagas.apply(lambda row_vagas: dicionario_escolas.get(row_vagas["CR Curso"], ""), axis=1)

# Criação das colunas de Análise (se existem vagas) e diferença (vagas totais - alunos cadastrados)
df_vagas.insert(11, 'Análise', "")
df_vagas.insert(12, 'Diferença', "")

df_vagas['Diferença'] = df_vagas.apply(
    lambda row_vagas: row_vagas["Nr Vagas Cadastradas"] - row_vagas["Qtde de Alunos Matri."], axis=1)
df_vagas['Análise'] = df_vagas.apply(lambda row_vagas: analisar_diferenca(row_vagas["Diferença"]), axis=1)

# Criação da coluna 'Tipo' (de turma)
df_vagas.insert(6, 'Tipo', "", True)
df_vagas['Tipo'] = df_vagas.apply(
    lambda row_vagas: dicionario_tipos.get(row_vagas["Agrupamento"], row_vagas["Agrupamento"]), axis=1)

# Remoção de linhas duplicadas (não essencial para este contexto)
df_vagas = df_vagas.drop_duplicates()

# Seleção da data atual
utc_now = pytz.utc.localize(datetime.datetime.utcnow())
dt_now = utc_now.astimezone(pytz.timezone("America/Sao_Paulo"))

# Criação do arquivo de saída
workbook = xlsxwriter.Workbook('arquivo_vagas ' + str(dt_now.strftime("%Y-%m-%d %H-%M")) + '.xlsx')
worksheet = workbook.add_worksheet()

# Criação de estilos de células
formato_verde = workbook.add_format({'font_color': '#006100'})
formato_verde.set_bg_color('#C6EFCE')
formato_amarelo = workbook.add_format({'font_color': '#9C5700'})
formato_amarelo.set_bg_color('#FFEB9C')
formato_vermelho = workbook.add_format({'font_color': '#9C0031'})
formato_vermelho.set_bg_color('#FFC7CE')

# Looping para impressão dos resultados
row = 1
for index, linha in df_vagas.iterrows():
    disciplina = linha[8]
    if disciplina.__contains__("Eu, Não Robô") or disciplina.__contains__("Eu, não Robô"):
        linha[6] = "Slash"  # Tipo de turma

    # Para cada coluna, exceto 'Data da Semana'
    for x in range(0, 15):
        item = str(linha[x])
        if item == "nan":
            item = ''
        if x == 12:  # Coluna de análise
            if item == "Vagas Disponíveis":
                worksheet.write(row, x, item, formato_verde)
            elif item == "Não há vagas":
                worksheet.write(row, x, item, formato_amarelo)
            elif item == "Vagas excedidas":
                worksheet.write(row, x, item, formato_vermelho)
            else:
                worksheet.write(row, x, item)
        elif x == 13:  # Coluna de diferença
            if str(linha[x - 1]) == "Vagas Disponíveis":
                worksheet.write(row, x, item.split('.')[0], formato_verde)
            elif str(linha[x - 1]) == "Não há vagas":
                worksheet.write(row, x, item.split('.')[0], formato_amarelo)
            elif str(linha[x - 1]) == "Vagas excedidas":
                worksheet.write(row, x, item.split('.')[0], formato_vermelho)
            else:
                worksheet.write(row, x, item.split('.')[0])
        else:
            worksheet.write(row, x, item)

    row += 1

# Formatação da tabela (cabeçalho)
worksheet.add_table(0, 0, row - 1, 14, {'style': 'Table Style Light 1', 'columns': [
    {'header': 'Estabelecimento'},
    {'header': 'Escola'},
    {'header': 'CR Curso'},
    {'header': 'Curso'},
    {'header': 'Período'},
    {'header': 'Sigla'},
    {'header': 'Tipo'},
    {'header': 'Turma'},
    {'header': 'Disciplina'},
    {'header': 'Divisão'},
    {'header': 'Matriculados'},
    {'header': 'Vagas Cadastradas'},
    {'header': 'Análise'},
    {'header': 'Diferença'},
    {'header': 'Professor'}
]})

# Formatação da tabela (largura das colunas)
worksheet.set_column('A:A', 18)
worksheet.set_column('B:B', 23)
worksheet.set_column('C:C', 10.5)
worksheet.set_column('D:D', 50)
worksheet.set_column('E:E', 10)
worksheet.set_column('F:F', 15)
worksheet.set_column('G:G', 11)
worksheet.set_column('H:H', 43)
worksheet.set_column('I:I', 65)
worksheet.set_column('J:J', 9)
worksheet.set_column('K:K', 28)
worksheet.set_column('L:L', 22)
worksheet.set_column('M:M', 17)
worksheet.set_column('N:N', 12)
worksheet.set_column('O:O', 51)

workbook.close()
