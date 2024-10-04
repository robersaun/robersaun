"""Conversor do Excel do aSc para o Template PowerCubus

Script para gerar o template de importação do PowerCubus com os dados extraídos do aSc.
Recebe como entrada o arquivo excel gerado pelo script 'asc_xml_para_excel.py'.
Seleciona todos os cartões do aSc, mesmo os não alocados.

Não é mais utilizado, versão descontinuada (utilizar 'asc_xlsx_para_powercubus_fixas.py').

Desenvolvido por Vinicius Tozo
Última atualização: 06/01/2022
"""
import json
import math
import random
import re
import string
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas


def encontra_etapa(turma):
    try:
        etapa = turma.split(';')[2]
    except IndexError:
        etapa = turma
    return etapa


def encontra_curso(turma, dicionario):
    try:
        curso = dicionario.get(turma.split(';')[1], turma.split(';')[1])
    except IndexError:
        curso = turma
    return curso


def converte_turma(turma, dicionario):
    try:
        campus, curso, periodo, letra, turno, cr, *outros = turma.split(';')
    except ValueError:
        print(f'Não foi possível converter o formato da turma {turma}')
        return turma

    if curso not in dicionario:
        print('Sigla não encontrada para o curso: ' + curso)

    resultado = cr + ' - ' + dicionario.get(curso, curso) + ' - ' + periodo + letra + ' - ' + turno

    # Semipresencial possui a informação do módulo após os dados da turma
    if outros:
        modulo = outros[0]
        modulo = modulo.replace('1º MÓDULO', ' (1°M)')
        modulo = modulo.replace('2º MÓDULO', ' (2°M)')

        resultado += modulo

    return resultado


def cria_sigla_disciplina(nome, dicionario):
    # Converte numerais romanos no final do nome
    nome = re.sub(' I$', ' 1', nome)
    nome = re.sub(' II$', ' 2', nome)
    nome = re.sub(' III$', ' 3', nome)
    nome = re.sub(' IV$', ' 4', nome)
    nome = re.sub(' V$', ' 5', nome)
    nome = re.sub(' VI$', ' 6', nome)
    nome = re.sub(' VII$', ' 7', nome)
    nome = re.sub(' VIII$', ' 8', nome)

    # Converte todos os whitespaces para um espaço
    # Serve para remover múltiplos espaços seguidos
    nome = nome.replace('-', ' ')
    nome = ' '.join(nome.split())

    # Se o nome da disciplina já possui menos de 10 caracteres, utiliza o nome como sigla
    if len(nome) < 10:
        return nome

    # Remove palavras que começam com lowercase
    palavras = ' '.join(filter(lambda x: not x[0].islower(), nome.strip().split(' ')))
    palavras = palavras.split(' ')

    # Pega as iniciais de cada palavra e o máximo possível da primeira palavra sem ultrapassar o limite de 10 caracteres
    # Exemplo: 'Leitura e Escrita de Textos Técnico-Científicos' fica 'Leitu ETTC'
    sigla = ''
    for palavra in palavras[1:6]:
        sigla += palavra[0]
    sigla = palavras[0][0:9 - len(sigla)] + ' ' + sigla

    sigla = sigla.upper()

    while sigla in dicionario.values():
        sigla = sigla[:-1] + random.choice(string.ascii_uppercase)

    return sigla


def cria_dicionario_grupos(df_groups):
    df_groups['groups_studentcount'].fillna('0', inplace=True)
    dicionario = {}
    for index, row in df_groups.iterrows():
        if row['groups_id'] not in dicionario:
            name = row['groups_name'].replace('Turma completa', 'TC')
            studentcount = str(int(row['groups_studentcount'], ))
            dicionario[row['groups_id']] = f"{name} ({studentcount})"

    return dicionario


def cria_dicionario_turmas(df_classes, dict_siglas):
    dicionario = {}
    for index, row in df_classes.iterrows():
        if row['classes_id'] not in dicionario:
            dicionario[row['classes_id']] = converte_turma(row['classes_short'], dict_siglas)
    return dicionario


def cria_abreviacao_professor(nome, dict_abreviacao):
    if nome == 'SEM PROFESSOR':
        return f'S. Prof ' \
               f'{random.choice(string.ascii_uppercase + string.digits)}' \
               f'{random.choice(string.ascii_uppercase + string.digits)}'

    if nome in dict_abreviacao:
        return dict_abreviacao[nome]
    else:
        abreviacao = nome.split(' ')[0][0:8] + ' ' + nome.split(' ')[-1][0]

        while abreviacao in dict_abreviacao.values():
            abreviacao = abreviacao[:-1] + random.choice(string.ascii_uppercase)

        dict_abreviacao[nome] = abreviacao
        return abreviacao


def cria_codigo_aula(row):
    # Variável para indicar se a turma é a turma base
    e_turma_base = '0'
    if (row['Turma'] == row['Turma Base']) and (row['Grupo'] == row['Grupo Base']):
        e_turma_base = '1'

    #         ID_Aula  ; ID_Agrupamento               ; Divisao      ; É_Turma_Base ; É_Horario_Flutuante
    return f"{row.name};{row['Código do agrupamento']};{row['Grupo']};{e_turma_base};{row['Horário Flutuante']}"


def cria_dicionario_aulas_flutuantes(df_cards):
    dicionario = {}
    for index, row in df_cards.iterrows():
        if row['cards_lessonid'] not in dicionario:
            if row['cards_days'] == 1:
                dicionario[row['cards_lessonid']] = '1'
            else:
                dicionario[row['cards_lessonid']] = '0'

    return dicionario


def cria_dicionario_aulas_fixas(df_cards, df_daysdefs, df_periods):

    # Substitui valores com vírgula por zero, para tratar as definições de 'Todos os dias'
    df_daysdefs.loc[df_daysdefs['daysdefs_days'].str.contains(','), 'daysdefs_days'] = 0

    # Converte todos os ids para string antes de juntar os dados
    df_cards['cards_days'] = df_cards['cards_days'].astype(int)
    df_cards['cards_period'] = df_cards['cards_period'].astype(str)
    df_daysdefs['daysdefs_days'] = df_daysdefs['daysdefs_days'].astype(int)
    df_periods['periods_period'] = df_periods['periods_period'].astype(str)

    df_cards = pandas.merge(df_cards, df_daysdefs, left_on='cards_days', right_on='daysdefs_days', how='left')
    df_cards = pandas.merge(df_cards, df_periods, left_on='cards_period', right_on='periods_period', how='left')

    dicionario = {}
    for index, row in df_cards.iterrows():
        if row['cards_lessonid'] not in dicionario:
            dicionario[row['cards_lessonid']] = ''
        dicionario[row['cards_lessonid']] += (row['daysdefs_short'] + ' ' + row['periods_starttime'] + ',')

    for index in dicionario:
        dicionario[index] = dicionario[index].strip(',')

    return dicionario


def verifica_caracteres(df_lessons):
    limite = [
        30, 10, 20, 100, 10, 50, 100, 10, 100, 50, 10, 10, 10, 10, 10, 50, 50, 50, 50, 50, 50, 50, 50, 10, 10, 50, 10,
        999
    ]

    for i in range(len(df_lessons.columns)):
        maior_valor = df_lessons[df_lessons.columns[i]].astype(str).str.len().max()
        if maior_valor > limite[i]:
            print(f'Limite de caracteres atingido na coluna {df_lessons.columns[i]} ({maior_valor}/{limite[i]})')


def trata_local_aula(nome):
    nome = str(nome)
    if nome and ';' in nome and 'Bloco ' in nome:
        bloco, sala, *outros = nome.split(';')
        bloco = bloco.replace('Bloco ', 'BL').split('-')[0].strip()
        nome = f'{bloco};{sala}'
    return nome


def corrige_divisoes_grupos(df_lessons):
    agrupamentos = {}
    for i, row in df_lessons.iterrows():

        if not row['Horário Fixo']:
            continue

        key = row['Turma'] + row['Horário Fixo']
        value = row['Código do agrupamento']

        if key not in agrupamentos:
            agrupamentos[key] = value
        elif value != agrupamentos[key]:
            df_lessons.at[i, 'Código do agrupamento'] = agrupamentos[key]

    return df_lessons


def main():
    # Seleciona o arquivo do aSc Timetables
    print('Selecione o arquivo excel')
    Tk().withdraw()
    file_name_xlsx = askopenfilename(filetypes=[('xlsx', '.xlsx')])

    # Lê o arquivo e grava em um dataframe
    df_lessons = pandas.read_excel(file_name_xlsx, 'lessons')
    df_classes = pandas.read_excel(file_name_xlsx, 'classes')
    df_teachers = pandas.read_excel(file_name_xlsx, 'teachers')
    df_subjects = pandas.read_excel(file_name_xlsx, 'subjects')
    df_classrooms = pandas.read_excel(file_name_xlsx, 'classrooms')
    df_groups = pandas.read_excel(file_name_xlsx, 'groups')
    df_cards = pandas.read_excel(file_name_xlsx, 'cards')
    df_periods = pandas.read_excel(file_name_xlsx, 'periods')
    df_daysdefs = pandas.read_excel(file_name_xlsx, 'daysdefs')

    # Renomeia as colunas para adicionar um prefixo
    df_lessons.columns = 'lessons_' + df_lessons.columns.values
    df_classes.columns = 'classes_' + df_classes.columns.values
    df_teachers.columns = 'teachers_' + df_teachers.columns.values
    df_subjects.columns = 'subjects_' + df_subjects.columns.values
    df_classrooms.columns = 'classrooms_' + df_classrooms.columns.values
    df_groups.columns = 'groups_' + df_groups.columns.values
    df_cards.columns = 'cards_' + df_cards.columns.values
    df_periods.columns = 'periods_' + df_periods.columns.values
    df_daysdefs.columns = 'daysdefs_' + df_daysdefs.columns.values

    # Cria os dicionários que serão utilizados
    dict_siglas = json.load(open('siglas.json', encoding='utf-8'))
    dicionario_grupos = cria_dicionario_grupos(df_groups)
    dicionario_turmas = cria_dicionario_turmas(df_classes, dict_siglas)

    # Adiciona coluna indicando aula em horário flutuante
    # Valor default do get é 0 para que as aulas sem cartão alocado não fiquem como horário flutuante
    dicionario_aulas_flutuantes = cria_dicionario_aulas_flutuantes(df_cards)
    df_lessons['Horário Flutuante'] = df_lessons.apply(
        lambda row: dicionario_aulas_flutuantes.get(row['lessons_id'], '0'), axis=1)

    # Adiciona coluna indicando aula em horário fixo
    dicionario_aulas_fixas = cria_dicionario_aulas_fixas(df_cards, df_daysdefs, df_periods)
    df_lessons['Horário Fixo'] = df_lessons.apply(
        lambda row: dicionario_aulas_fixas.get(row['lessons_id'], ''), axis=1)

    # Separa aulas compartilhadas em linhas separadas para cada turma + grupo
    df_lessons['lessons_classids'] = df_lessons['lessons_classids'].astype(str).str.split(',')
    df_lessons['lessons_groupids'] = df_lessons['lessons_groupids'].astype(str).str.split(',')
    df_lessons.insert(0, 'Turma Base', '', True)
    for i, value in df_lessons.iterrows():
        new_value = []
        for j in range(len(value['lessons_classids'])):
            new_value.append(value['lessons_classids'][j] + ' ' + value['lessons_groupids'][j])
        df_lessons.at[i, 'lessons_classids'] = new_value

        # A primeira turma fica salva como a turma base
        df_lessons.at[i, 'Turma Base'] = dicionario_turmas.get(new_value[0].split(' ')[0], new_value[0].split(' ')[0])
        df_lessons.at[i, 'Grupo Base'] = dicionario_grupos.get(new_value[0].split(' ')[1], new_value[0].split(' ')[1])

    df_lessons = df_lessons.explode('lessons_classids')

    # Separa os valores novamente
    df_lessons['lessons_groupids'] = df_lessons.apply(lambda row: row['lessons_classids'].split(' ')[1], axis=1)
    df_lessons['lessons_classids'] = df_lessons.apply(lambda row: row['lessons_classids'].split(' ')[0], axis=1)

    # Separa aulas com mais de um professor em multiplas linhas
    df_lessons['lessons_teacherids'] = df_lessons['lessons_teacherids'].str.split(',')
    df_lessons = df_lessons.explode('lessons_teacherids')

    # Converte ids para mesmo tipo
    df_lessons['lessons_classids'] = df_lessons['lessons_classids'].astype(str)
    df_classes['classes_id'] = df_classes['classes_id'].astype(str)
    df_lessons['lessons_teacherids'] = df_lessons['lessons_teacherids'].astype(str)
    df_teachers['teachers_id'] = df_teachers['teachers_id'].astype(str)
    df_lessons['lessons_classroomids'] = df_lessons['lessons_classroomids'].astype(str)
    df_classrooms['classrooms_id'] = df_classrooms['classrooms_id'].astype(str)

    # Junta todos os dados em um único dataframe
    df_lessons = pandas.merge(df_lessons, df_classes, left_on='lessons_classids', right_on='classes_id')
    df_lessons = pandas.merge(df_lessons, df_teachers, how='left', left_on='lessons_teacherids', right_on='teachers_id')
    df_lessons = pandas.merge(df_lessons, df_subjects, left_on='lessons_subjectid', right_on='subjects_id')
    df_lessons = pandas.merge(
        df_lessons, df_classrooms, how='left', left_on='lessons_classroomids', right_on='classrooms_id')

    # df_lessons = pandas.merge(df_lessons, df_groups, how='left', left_on='lessons_groupids', right_on='groups_id')
    df_lessons['Grupo'] = df_lessons.apply(lambda row: dicionario_grupos.get(row['lessons_groupids']), axis=1)

    # Renomeia as colunas
    df_lessons.rename(columns={
        'lessons_id': 'Código do agrupamento',
        'lessons_periodsperweek': 'Aulas',
        'classes_short': 'Turma',
        'teachers_short': 'Nome do professor',
        'subjects_short': 'Nome da disciplina',
        'classrooms_name': 'Local de aula',
    }, inplace=True)

    # Insere as colunas do template
    df_lessons.insert(0, 'Etapa', '', True)
    df_lessons.insert(0, 'Curso', '', True)
    df_lessons.insert(0, 'Área', '-', True)
    df_lessons.insert(0, 'Sigla da disciplina', '', True)
    df_lessons.insert(0, 'Nome abreviado do professor', '', True)
    df_lessons.insert(0, 'Email do professor', '', True)
    df_lessons.insert(0, 'Máximo de aulas diárias', '', True)
    df_lessons.insert(0, 'Agrupamento de aulas', '1', True)
    df_lessons.insert(0, 'Máximo de dias de aula na semana', '', True)
    df_lessons.insert(0, 'Permitir aulas em dias consecutivos', '1', True)
    df_lessons.insert(0, 'Código de origem da turma', '', True)
    df_lessons.insert(0, 'Código de origem da disciplina', df_lessons['subjects_partner_id'], True)
    df_lessons.insert(0, 'Código de origem do professor', df_lessons['teachers_customfield1'], True)
    df_lessons.insert(0, 'Código de origem do local', '', True)
    df_lessons.insert(0, 'Código de origem da área', '', True)
    df_lessons.insert(0, 'Nome da unidade da turma', 'Curitiba', True)
    df_lessons.insert(0, 'Grupo de períodos', 'Graduação Presencial', True)
    df_lessons.insert(0, 'Nome da unidade do local', 'Curitiba', True)
    df_lessons.insert(0, 'Sigla da unidade da turma', 'CWB', True)
    df_lessons.insert(0, 'Sigla da unidade do local', 'CWB', True)
    df_lessons.insert(0, 'Código de origem da aula', '', True)

    # Preenche os dados nas colunas inseridas
    df_lessons['Etapa'] = df_lessons.apply(lambda row: encontra_etapa(row['Turma']), axis=1)
    df_lessons['Curso'] = df_lessons.apply(lambda row: encontra_curso(row['Turma'], dict_siglas), axis=1)
    df_lessons['Turma'] = df_lessons.apply(lambda row: converte_turma(row['Turma'], dict_siglas), axis=1)

    # Trata a coluna do nome do professor
    df_lessons['Nome do professor'].fillna(';Sem Professor', inplace=True)
    df_lessons['Nome do professor'] = df_lessons.apply(lambda row: row['Nome do professor'].split(';')[-1].upper(),
                                                       axis=1)

    # Trata a coluna do máximo de dias de aula na semana
    df_lessons['Máximo de dias de aula na semana'] = df_lessons.apply(
        lambda row: int(math.ceil((int(row['Aulas']) / 6))), axis=1)

    # Deixa o nome da disciplina em letras maiusculas para normalizar os dados
    df_lessons['Nome da disciplina'] = df_lessons['Nome da disciplina'].str.upper()

    # Gera uma sigla única para cada disciplina
    dict_siglas = {}
    for index, valor in df_lessons.iterrows():
        nome_disciplina = valor['Nome da disciplina']
        if nome_disciplina not in dict_siglas:
            dict_siglas[nome_disciplina] = cria_sigla_disciplina(nome_disciplina, dict_siglas)
    df_lessons['Sigla da disciplina'] = df_lessons.apply(lambda row: dict_siglas.get(row['Nome da disciplina']), axis=1)

    # Gera um id sequencial para cada agrupamento
    agrupamentos = []
    for index, valor in df_lessons.iterrows():
        id_asc = valor['Código do agrupamento']
        if id_asc not in agrupamentos:
            agrupamentos.append(id_asc)
    df_lessons['Código do agrupamento'] = df_lessons.apply(
        lambda row: agrupamentos.index(row['Código do agrupamento']) + 1, axis=1)

    # Abreviação do professor até 10 caracteres e único pra cada professor
    dict_abreviacao = {}
    df_lessons['Nome abreviado do professor'] = df_lessons.apply(
        lambda row: cria_abreviacao_professor(row['Nome do professor'], dict_abreviacao), axis=1)

    df_lessons['Local de aula'].fillna('Sala não informada', inplace=True)
    df_lessons['Local de aula'] = df_lessons.apply(
        lambda row: f"Sala não informada {row['Grupo'][0]}" if row['Local de aula'] == 'Sala não informada'
        else row['Local de aula'], axis=1)
    df_lessons['Local de aula'] = df_lessons.apply(lambda row: trata_local_aula(row['Local de aula']), axis=1)

    # Corrige múltiplos grupos de uma mesma turma com aulas ao mesmo tempo
    df_lessons = corrige_divisoes_grupos(df_lessons)

    # Preenche o código de origem da aula
    df_lessons['Código de origem da aula'] = df_lessons.apply(lambda row: cria_codigo_aula(row), axis=1)

    # Filtra apenas as colunas que devem ser exibidas
    df_lessons = df_lessons[[
        'Turma',
        'Etapa',
        'Curso',
        'Nome da disciplina',
        'Sigla da disciplina',
        'Área',
        'Nome do professor',
        'Nome abreviado do professor',
        'Email do professor',
        'Local de aula',
        'Aulas',
        'Máximo de aulas diárias',
        'Agrupamento de aulas',
        'Máximo de dias de aula na semana',
        'Permitir aulas em dias consecutivos',
        'Código de origem da turma',
        'Código de origem da disciplina',
        'Código de origem do professor',
        'Código de origem do local',
        'Código de origem da área',
        'Nome da unidade da turma',
        'Grupo de períodos',
        'Nome da unidade do local',
        'Sigla da unidade da turma',
        'Sigla da unidade do local',
        'Código de origem da aula',
        'Código do agrupamento',
        'Horário Fixo',
    ]]

    # Ordena por agrupamento
    df_lessons.sort_values('Código do agrupamento')

    # Verifica número de caracteres por coluna
    verifica_caracteres(df_lessons)

    # Salva o resultado em um arquivo excel
    df_lessons.to_excel(file_name_xlsx.replace('.xlsx', ' PowerCubus.xlsx'), sheet_name='Template', index=False)


if __name__ == '__main__':
    main()
