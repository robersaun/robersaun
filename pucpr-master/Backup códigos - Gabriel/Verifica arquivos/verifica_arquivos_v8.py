# -*- coding: Windows-1252 -*-
import os
from tkinter import Tk, filedialog
from tkinter.filedialog import askopenfilename
import pandas as pd


def data_admissao_f(ano_mes_admissao):
    append = ''
    if '-' in ano_mes_admissao:
        ano_admissao = int(ano_mes_admissao.split('-')[0])
        mes_admissao = str(ano_mes_admissao.split('-')[1])
        if ano_mes_admissao == '2015-1':
            append = 'Não foi encontrado'
        else:
            append = f'{ano_mes_admissao.split("-")[2]}/{mes_admissao}/{ano_admissao}'
    else:
        mes_admissao = int(ano_mes_admissao.split('/')[1])
        ano_admissao = int(ano_mes_admissao.split('/')[2])
        append = ano_mes_admissao
    return [append, ano_admissao, mes_admissao]


def lista_matrizes(root):

    lista_profs, admissao, lista_pendencias, lista_doc = [], [], [], []

    # Seleção do arquivo
    Tk().withdraw()
    planilha = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de admissões')

    df = pd.read_excel(planilha, dtype=str, sheet_name=0, header=3, usecols='C:D')
    df = df.fillna("")

    prof_relatorio, prof_r_admissao = df['NOME'].tolist(), df['ADMISSÃO'].tolist()
    prof_relatorio.remove('')

    # Para cada diretório dentro do diretório raiz
    for dirpath, dirnames, filenames in os.walk(root):

        # Se o diretório possuir outros diretórios dentro dele, ignora
        if dirnames:
            continue

        # Verifica o nome da pasta, para pegar o nome do prof, e corrrige a fonte para maiúscula
        prof = str(dirpath.replace('\\', '/').split('/')[-2]).upper()
        if len(prof) == 1:
            prof = str(dirpath.replace('\\', '/').split('/')[-1]).upper()
        # Verifica o nome da pasta, para tratar as pastas que não tem nome de prof
        if (len(prof) <= 1 or prof == 'ANTIGO' or 'HIPS 20' in prof or
                'NOVA PASTA' in prof or 'TERMO' in prof or 'ADMISSÃO' in prof):
            continue
        # Busca o prof e a data de admissão
        try:
            index = df.index[df['NOME'] == prof].tolist()[0]
            ano_mes_admissao = str(df['ADMISSÃO'][index].split()[0])
        except:
            ano_mes_admissao = '2015-1'

        # Inclui a data de admissão para identificar os documentos faltantes
        da = data_admissao_f(ano_mes_admissao)

        admissao.append(da[0])
        ano_admissao = da[1]
        mes_admissao = da[2]

        lista_profs.append(prof)
        ano_periodo, ano_periodo_faltante = [], []

        # Verifica se o ano de admissão do prof, para não cobrar os documentos desnecessários
        if ano_admissao > 2015:
            ano_verifica = ano_admissao
            if 1 <= int(mes_admissao) < 6:
                semestre = 1
            else:
                semestre = 2
        else:
            ano_verifica = 2015
            # Como quem entra no else já estava admitido previamente, o semestre fica como 1 independente do mês
            semestre = 1

        docs_existentes = ''
        doc_periodos = []
        for file in filenames:
            if file.endswith('.pdf') or file.endswith('.PDF'):
                file_year = file.replace("-", ".").split('20')[-1].split('.')
                if file_year[0] == '':
                    file_year[0] = '20'
                file_year = f'20{file_year[0]}.{file_year[1]}'.split(' ')[0].split('-')[0].split('_')[0]
                ano_periodo.append(file.replace("-", "."))
                doc_periodos.append(file_year)

        if len(doc_periodos) > 0:
            doc_periodos = sorted(list(set(doc_periodos)))
        for doc in doc_periodos:
            if doc == doc_periodos[-1]:
                docs_existentes += doc
            else:
                docs_existentes += f'{doc}, '
        lista_doc.append(docs_existentes)

        # Verifica quais anos não estão presentes
        while ano_verifica <= 2021:
            if (str(ano_verifica) + '.2' == str(ano_admissao) + '.' + str(semestre) and
                    str(ano_verifica) + '.2' not in str(ano_periodo)):
                ano_periodo_faltante.append(f'{ano_verifica}.2')
            else:
                if str(ano_verifica) + '.1' not in str(ano_periodo):
                    ano_periodo_faltante.append(f'{ano_verifica}.1')
                if str(ano_verifica) + '.2' not in str(ano_periodo) and ano_verifica != 2021:
                    ano_periodo_faltante.append(f'{ano_verifica}.2')
            ano_verifica += 1

        # Remove os valores duplicados e ordena a lista
        if len(ano_periodo_faltante) > 0:
            ano_periodo_faltante = sorted(list(set(ano_periodo_faltante)))
        ano_periodo_pendente = ''
        for ap in ano_periodo_faltante:
            if ap == ano_periodo_faltante[-1]:
                ano_periodo_pendente += f'{ap}'
                break
            ano_periodo_pendente += f'{ap}, '
        lista_pendencias.append(ano_periodo_pendente)

    for nome_prof in range(len(prof_relatorio)):
        if prof_relatorio[nome_prof] not in lista_profs:
            lista_profs.append(prof_relatorio[nome_prof])
            lista_pendencias.append('Não há pasta de documentos')
            data_admissao = str(prof_r_admissao[nome_prof].split()[0])
            # Ajusta o formato da data
            if '-' in data_admissao:
                admissao.append(f'{data_admissao.split("-")[2]}/'
                                f'{data_admissao.split("-")[1]}/'
                                f'{data_admissao.split("-")[0]}')
            else:
                admissao.append(data_admissao)
            lista_doc.append('Não encontrado')

    dados = {'Docente': lista_profs,
             'Admissão': admissao,
             'Documentos TACH Pendentes': lista_pendencias,
             'Documentos TACH Existêntes': lista_doc}
    ndf = pd.DataFrame(data=dados)
    ndf.to_excel('Arquivos_pendentes_profs-v8.xlsx', index=False, sheet_name='Pendências')


if __name__ == '__main__':
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    print()
    lista_matrizes(pasta_raiz)
