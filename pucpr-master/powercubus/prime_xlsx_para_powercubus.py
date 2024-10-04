"""Conversor do Excel do Prime para o Template PowerCubus

Script para gerar o template de importação do PowerCubus com os dados exportados do Prime.
Recebe como entrada o arquivo excel gerado pelo script 'asc_xml_para_excel.py'.
Gera um arquivo excel para cada tipo de dado, seguindo o modelo de importação do PowerCubus

Utilizado pelo analista de TI no processo de elaboração de horários no PowerCubus.

Desenvolvido por Vinicius Tozo
Última atualização: 21/02/2022
"""
import json
import random
import re
import string
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas


def cria_sigla_disciplina(nome, dicionario):
    # Remove caracteres especiais
    nome = ''.join([i for i in nome if i.isdigit() or i.isalpha() or i.isspace()])

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
    nome = nome.strip()

    sigla = ''

    # Remove palavras que começam com lowercase
    if ' ' in nome:
        palavras = ' '.join(filter(lambda x: not x[0].islower(), nome.strip().split(' ')))
        palavras = palavras.split(' ')

        # Pega as iniciais de cada palavra e o máximo possível da primeira palavra sem passar o limite de caracteres
        # Exemplo: 'Leitura e Escrita de Textos Técnico-Científicos' fica 'Leitu ETTC'
        for palavra in palavras[1:6]:
            sigla += palavra[0]
        sigla = palavras[0][0:9 - len(sigla)] + ' ' + sigla
    else:
        sigla = nome
        if len(nome) > 10:
            sigla = nome[0:10]

    sigla = sigla.upper()

    while sigla in dicionario.values():
        sigla = sigla[:-1] + random.choice(string.ascii_uppercase)

    return sigla


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


def cria_abreviacao_professor(nome, dict_abreviacao):
    nome = nome.strip()
    if nome in dict_abreviacao:
        return dict_abreviacao[nome]
    else:
        split = nome.split(' ')
        if len(split) >= 3:
            # Se tiver pelo menos dois sobrenomes pega a inicial do penúltimo e do último
            abreviacao = split[0][0:7] + ' ' + split[-2][0] + split[-1][0]
        else:
            # Se não pega a inicial apenas do último
            abreviacao = split[0][0:8] + ' ' + split[-1][0]

        while abreviacao in dict_abreviacao.values():
            abreviacao = abreviacao[:-1] + random.choice(string.ascii_uppercase)

        dict_abreviacao[nome] = abreviacao
        return abreviacao


def create_template_classes(df, file_name_xlsx):
    # Lê o arquivo contendo a sigla de cada curso, para reduzir o tamanho dos nomes
    dict_siglas = json.load(open('siglas.json', encoding='utf-8'))

    # Cria as colunas do template e preenche com os dados necessários
    df['Turma'] = df.apply(lambda row: converte_turma(row['short'], dict_siglas), axis=1)
    df['Etapa'] = df.apply(lambda row: row['short'].split(';')[2], axis=1)
    df['Curso'] = df.apply(lambda row: row['Turma'].split(' - ')[1], axis=1)
    df['Local de aula padrão'] = df.apply(lambda row: '', axis=1)
    df['Código de origem'] = df.apply(lambda row: row['partner_id'], axis=1)
    df['Grupo de períodos'] = df.apply(lambda row: 'Graduação Presencial', axis=1)
    df['Unidade do local de aula padrão'] = df.apply(lambda row: row['short'].split(';')[0], axis=1)
    df['Sigla da unidade do local'] = df.apply(lambda row: row['short'].split(';')[0], axis=1)
    df['Unidade da turma'] = df.apply(lambda row: row['short'].split(';')[0], axis=1)
    df['Sigla da unidade da turma'] = df.apply(lambda row: row['short'].split(';')[0], axis=1)

    # Remove colunas não utilizadas e ordena as restantes
    df = df[[
        'Turma',
        'Etapa',
        'Curso',
        'Local de aula padrão',
        'Código de origem',
        'Grupo de períodos',
        'Unidade do local de aula padrão',
        'Sigla da unidade do local',
        'Unidade da turma',
        'Sigla da unidade da turma',
    ]]

    # Verifica se existem registros duplicados e emite alerta
    duplicated = df[df.duplicated(['Turma'])]
    for k, v in duplicated.iterrows():
        print('Turma duplicada no Prime, corrija manualmente antes de importar: ', v['Turma'])

    # Salva o resultado como arquivo XLSX
    df.to_excel(file_name_xlsx.replace('.xlsx', ' Turmas PowerCubus.xlsx'), sheet_name='Template', index=False)


def create_template_teachers(df, file_name_xlsx):
    # Cria as colunas do template e preenche com os dados necessários
    df['Nome do professor'] = df.apply(lambda row: row['short'].upper().split(';')[1], axis=1)
    # Abreviação do professor até 10 caracteres e único pra cada professor
    dict_abreviacao = {}
    df['Nome abreviado do professor'] = \
        df.apply(lambda row: cria_abreviacao_professor(row['Nome do professor'], dict_abreviacao), axis=1)
    df['Email do professor'] = df.apply(lambda row: '', axis=1)
    df['Limite de aulas diárias'] = df.apply(lambda row: 8, axis=1)
    df['Horas-atividade'] = df.apply(lambda row: '', axis=1)
    df['Reduzir dias de trabalho'] = df.apply(lambda row: 1, axis=1)
    df['Reduzir janelas'] = df.apply(lambda row: 1, axis=1)
    df['Código de origem'] = df.apply(lambda row: row['id'], axis=1)

    # Remove colunas não utilizadas e ordena as restantes
    df = df[[
        'Nome do professor',
        'Nome abreviado do professor',
        'Email do professor',
        'Limite de aulas diárias',
        'Horas-atividade',
        'Reduzir dias de trabalho',
        'Reduzir janelas',
        'Código de origem',
    ]]

    # Verifica se existem registros duplicados e emite alerta
    duplicated = df[df.duplicated(['Nome do professor'])]
    for k, v in duplicated.iterrows():
        print('Professor duplicado no Prime, corrija manualmente antes de importar: ', v['Nome do professor'])

    # Salva o resultado como arquivo XLSX
    df.to_excel(file_name_xlsx.replace('.xlsx', ' Professores PowerCubus.xlsx'), sheet_name='Template', index=False)


def create_dict_subjects(names):
    # Seleciona os diferentes valores para nomes de disciplinas
    names = names.drop_duplicates()
    names = names.values.tolist()

    dict_subjects = {}
    for name in names:
        dict_subjects[name] = cria_sigla_disciplina(name, dict_subjects)
    return dict_subjects


def create_template_subjects(df, file_name_xlsx):
    # Cria as colunas do template e preenche com os dados necessários
    df['Nome da disciplina'] = df.apply(lambda row: row['short'].strip(), axis=1)
    dict_subjects = create_dict_subjects(df['Nome da disciplina'])
    df['Sigla da disciplina'] = df.apply(lambda row: dict_subjects.get(row['Nome da disciplina']), axis=1)
    df['Local de aula padrão'] = df.apply(lambda row: '', axis=1)
    df['Área de conhecimento'] = df.apply(lambda row: '', axis=1)
    df['Código de origem'] = df.apply(lambda row: '', axis=1)
    # TODO: Preparar para os campus fora de sede, identificar o campus através do cr
    df['Unidade do local de aula padrão'] = df.apply(lambda row: 'CWB', axis=1)
    df['Sigla da unidade do local'] = df.apply(lambda row: 'CWB', axis=1)

    # Remove colunas não utilizadas e ordena as restantes
    df = df[[
        'Nome da disciplina',
        'Sigla da disciplina',
        'Local de aula padrão',
        'Área de conhecimento',
        'Código de origem',
        'Unidade do local de aula padrão',
        'Sigla da unidade do local',
    ]]

    df = df.drop_duplicates()

    # Verifica se existem registros duplicados e emite alerta
    duplicated = df[df.duplicated(['Nome da disciplina'])]
    for k, v in duplicated.iterrows():
        print('Disciplina duplicada no Prime, corrija manualmente antes de importar: ', v['Nome da disciplina'])

    # Salva o resultado como arquivo XLSX
    df.to_excel(file_name_xlsx.replace('.xlsx', ' Disciplinas PowerCubus.xlsx'), sheet_name='Template', index=False)


def get_turno(start_time):
    if start_time < '12:40':
        # Manhã
        return 'M'
    elif '12:40' <= start_time < '18:15':
        # Tarde
        return 'T'
    else:
        # Noite
        return 'N'


def create_template_periods(df, file_name_xlsx):
    # Cria as colunas do template e preenche com os dados necessários
    df['Início'] = df.apply(lambda row: row['starttime'], axis=1)
    df['Fim'] = df.apply(lambda row: row['endtime'], axis=1)
    df['Turno'] = df.apply(lambda row: get_turno(row['starttime']), axis=1)
    df['Código de origem'] = df.apply(lambda row: '', axis=1)
    df['Grupo'] = df.apply(lambda row: 'Graduação Presencial', axis=1)

    # Remove colunas não utilizadas e ordena as restantes
    df = df[[
        'Início',
        'Fim',
        'Turno',
        'Código de origem',
        'Grupo',
    ]]

    # Verifica se existem registros duplicados e emite alerta
    duplicated = df[df.duplicated(['Início'])]
    for k, v in duplicated.iterrows():
        print('Período duplicado no Prime, corrija manualmente antes de importar: ', v['Início'])

    # Salva o resultado como arquivo XLSX
    df.to_excel(file_name_xlsx.replace('.xlsx', ' Periodos PowerCubus.xlsx'), sheet_name='Template', index=False)


def main():
    # Seleciona o arquivo exportado do Prime e convertido para XLSX
    print('Selecione o arquivo excel')
    Tk().withdraw()
    file_name_xlsx = askopenfilename(filetypes=[('xlsx', '.xlsx')])

    # Lê o arquivo e grava em um dataframe para cada tipo de dado
    df_classes = pandas.read_excel(file_name_xlsx, 'classes')
    df_teachers = pandas.read_excel(file_name_xlsx, 'teachers')
    df_subjects = pandas.read_excel(file_name_xlsx, 'subjects')
    df_periods = pandas.read_excel(file_name_xlsx, 'periods')

    # Cria o arquivo de de cada tipo de dado com base nos templates do PowerCubus
    create_template_classes(df_classes, file_name_xlsx)
    create_template_teachers(df_teachers, file_name_xlsx)
    create_template_subjects(df_subjects, file_name_xlsx)
    create_template_periods(df_periods, file_name_xlsx)
    # TODO: Validar número de caracteres


if __name__ == '__main__':
    main()
