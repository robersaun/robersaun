"""Gerador de Planilha Turmas

Script para gerar a 'Planilha Turmas' com base nas resoluções.
Recebe como entrada os seguintes arquivos no formato XLSX: template da planilha turmas, base de professores ativos,
dicionário de códigos de disciplina e arquivo unificado de resoluções.
Gera as 'Planilhas Turmas' com os dados selecionados, separadas por curso.
A planilha gerada vai conter as disciplinas apropriadas para cada período, de acordo com o ano indicado na matriz.

Pode ser utilizado pelo analista de TI para auxiliar no processo de elaboração de horários.


Desenvolvido por Vinicius Tozo
Última atualização: 19/08/2021
"""
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas


def cria_dicionario_disciplinas(arquivo):
    # Cria um dicionário com os códigos de disciplinas
    df_cod_disciplinas = pandas.read_excel(arquivo, engine='openpyxl')
    df_cod_disciplinas = df_cod_disciplinas[['Código', 'Nome']]
    dict_disciplina = {}
    for k, v in df_cod_disciplinas.iterrows():
        dict_disciplina[v['Nome']] = v['Código']
    return dict_disciplina


def cria_lista_professores(arquivo):
    # Cria uma lista com o nome dos professores
    df_professores = pandas.read_excel(arquivo, engine='openpyxl', skiprows=2)
    df_professores = df_professores[['NOME']]

    # Seleciona os diferentes valores para Escola
    df_professores = df_professores.drop_duplicates().sort_values(['NOME'])
    lista_professores = df_professores.values.tolist()

    # Transforma lista de listas em um lista única
    flat_list = []
    for sublist in lista_professores:
        for item in sublist:
            flat_list.append(item)

    return flat_list


def cria_df_listas(arquivo):
    # Cria uma lista com o nome dos professores
    df_listas = pandas.read_excel(arquivo, sheet_name='Listas', engine='openpyxl')
    return df_listas


def salva_planilha_turmas(dataframe, path, lista_professores, df_listas):
    # Cria a pasta de saída
    folders = '/'.join(path.split('/')[0:-1])
    if folders != '' and not os.path.exists(folders):
        os.makedirs(folders)

    # Converte o dataframe para um arquivo excel
    writer = pandas.ExcelWriter(path, engine='xlsxwriter')
    dataframe.to_excel(writer, sheet_name='Disciplinas Regulares', startrow=1, header=False, index=False)

    # Cria a aba de disciplinas regulares
    workbook = writer.book
    worksheet = writer.sheets['Disciplinas Regulares']

    # Descobre as dimensões do dataframe
    (max_row, max_col) = dataframe.shape

    # Adiciona uma tabela para formatação
    column_settings = [{'header': column} for column in dataframe.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {'style': 'Table Style Medium 17', 'columns': column_settings})

    # Formatação da tabela
    formato_cabecalho = workbook.add_format({'bg_color': '#595959', 'border': 1, 'font_color': 'white'})
    formato_prenchimento = workbook.add_format({'bg_color': '#C0504D', 'border': 1, 'font_color': 'white'})
    worksheet.set_column('A:A', 23)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 52)
    worksheet.set_column('D:D', 19)
    worksheet.set_column('E:E', 84)
    worksheet.set_column('F:F', 10)
    worksheet.set_column('G:G', 10)
    worksheet.set_column('H:H', 26)
    worksheet.set_column('I:I', 8)
    worksheet.set_column('J:J', 12)
    worksheet.set_column('K:K', 19)
    worksheet.set_column('L:L', 24)
    worksheet.set_column('M:M', 15)
    worksheet.set_column('N:N', 30)
    worksheet.set_column('O:Z', 50)

    # Escreve novamente o cabeçalho com a formatação correta
    for col_num, value in enumerate(dataframe.columns.values):
        if col_num >= 14 or col_num == 9:
            worksheet.write(0, col_num, value, formato_prenchimento)
        else:
            worksheet.write(0, col_num, value, formato_cabecalho)

    # Cria as demais abas do modelo
    ws_eletivas = workbook.add_worksheet('Eletivas Próprias')
    ws_especiais = workbook.add_worksheet('Turmas Especiais Previstas')
    ws_globais = workbook.add_worksheet('Global Classes')

    # Cria aba de professores
    df_professores = pandas.DataFrame(data=lista_professores, columns=['Nome do professor'])
    df_professores.to_excel(writer, sheet_name='Professores', startrow=1, header=False, index=False)
    worksheet = writer.sheets['Professores']

    # Descobre as dimensões do dataframe de professores
    (max_row, max_col) = df_professores.shape

    # Adiciona uma tabela para formatação da aba de professores
    column_settings = [{'header': column} for column in df_professores.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {'style': 'Table Style Medium 17', 'columns': column_settings})
    worksheet.set_column('A:A', 55)
    worksheet.write(0, 0, 'Nome do professor', formato_cabecalho)

    # Cria aba de listas
    df_listas.to_excel(writer, sheet_name='Listas', startrow=0, header=False, index=False)

    # Lista as colunas de cada aba
    eletivas = [
        'Escola/Campus', 'Curso', 'Cód. Do Curso', 'Período', 'Turma', 'Turno',
        'Demais períodos que podem cursar a disciplina', 'É a primeira oferta?', 'Código da Disciplina',
        'Disciplinas Eletivas de Oferta Própria', 'Número Total de Vagas', 'CH Teórica',
        'Nome completo do professor - Aulas Teóricas', 'Professor (Aulas Teóricas) é do grupo de risco?',
        'Retorno Presencial? (Aulas Teóricas)', 'Hy Flex? (Aulas Teóricas)', 'CH Prática',
        'Modulação da Prática', 'Nome completo do professor - Aulas Práticas',
        'Professor (Aulas Práticas) é do grupo de risco?', 'Retorno Presencial? (Aulas Práticas)',
        'Hy Flex? (Aulas Práticas)', 'CH Tutoria', 'Modulação Tutoria', 'Nome completo do professor - Tutoria',
        'Professor (Tutoria) é do grupo de risco?', 'Retorno Presencial? (Tutoria)', 'Hy Flex? (Tutoria)',
        'Qual o tipo de Ambiente? Sala ou  Laboratório?', 'Disciplina aberta ou fechada?',
        'Indicar se a disciplina será aberta para o Curso, para a Escola ou para Instituição',
        'Ofertar para outros cursos?', 'Se sim, ofertar para:', 'Informações adicionais']
    especiais = [
        'ESCOLA / CAMPUS', 'CURSO', 'CÓDIGO DO CURSO', 'TURNO OFERTA', 'PERÍODO OFERTA', 'TURMA',
        'CÓDIGO DA DISCIPLINA', 'NOME DISCIPLINA', 'PREVISÃO NÚMERO DE ALUNOS', 'CH TOTAL DA DISCIPLINA',
        'MODULAÇÃO TEÓRICA', 'QUANTIDADE AULAS TEÓRICAS', 'Selecione o Professor - Aulas Teóricas',
        'Professor (Aulas Teóricas) é do grupo de risco?', 'Retorno Presencial? (Aulas Teóricas)',
        'Hy Flex? (Aulas Teóricas)', 'MODULAÇÃO PRÁTICA', 'QUANTIDADE AULAS PRÁTICAS',
        'Selecione o Professor - Aulas Práticas', 'Professor (Aulas Práticas) é do grupo de risco?',
        'Retorno Presencial? (Aulas Práticas)', 'Hy Flex? (Aulas Práticas)', 'MODULAÇÃO TUTORIA',
        'QUANTIDADE AULAS TUTORIA', 'Selecione o Professor - Aulas Tutoria',
        'Professor (Tutoria) é do grupo de risco?', 'Retorno Presencial? (Tutoria)', 'Hy Flex? (Tutoria)',
        'Professor possui Carga Horária Não Letiva?', 'Indicar dia da semana - Dia 1',
        'Indicar Horário de Início - Dia 1', 'Indicar Horário de Término - Dia 1',
        'Teórico, Prático ou Tutorial - Dia 1',
        'Utiliza Sala Teórica ou Ambiente específico? (Especificar ambiente) - Dia 1',
        'Indicar dia da semana - Dia 2', 'Indicar Horário de Início - Dia 2',
        'Indicar Horário de Término - Dia 2', 'Teórico, Prático ou Tutorial - Dia 2',
        'Utiliza Sala Teórica ou Ambiente específico? (Especificar ambiente) - Dia 2',
        'Indicar dias, horários, modulação e ensalamento para disciplinas que acontecerão em mais de 2 dias da semana',
        'Observação']
    globais = [
        'ESCOLA', 'Código curso', 'Nome do curso', 'Período da disciplina',
        '"Tipo de disciplina (Obrigatória, optativa ou eletiva)"', '"Nível 1, 2 ou 3 (Global Classes)"', 'Turma',
        'A disciplina pode ser ofertada para outros curso?',
        'Código da disciplina em Português - indicar se é primeira oferta', 'X',
        'Nome da disciplina em português', 'Código da disciplina em Inglês', 'Nome da disciplina em inglês',
        'Professor', 'Nº de Vagas', 'Dia', 'Horário de início', 'Horário de fim', '"CH TOTAL DA DISCIPLINA"',
        '"Hora aula total da disciplina"', 'Observação', 'E-mail Institucional']

    # Transforma no formato do xlsxwriter
    eletivas = [{'header': column} for column in eletivas]
    especiais = [{'header': column} for column in especiais]
    globais = [{'header': column} for column in globais]

    # Adiciona tabelas
    ws_eletivas.add_table(0, 0, 10, len(eletivas) - 1, {'style': 'Table Style Medium 17', 'columns': eletivas})
    ws_especiais.add_table(0, 0, 10, len(especiais) - 1, {'style': 'Table Style Medium 17', 'columns': especiais})
    ws_globais.add_table(0, 0, 10, len(globais) - 1, {'style': 'Table Style Medium 17', 'columns': globais})

    # Salva o arquivo
    writer.save()


def identifica_tipo(at, ap):
    if at > 0 >= ap:
        return 'Teórico'
    elif ap > 0 >= at:
        return 'Prático'
    elif ap > 0 and at >= 0:
        return 'Ambos'
    else:
        return ''


def identifica_ch(row):
    if row['Tipo de Atividade'] == 'Teórico':
        return row['AT']
    elif row['Tipo de Atividade'] == 'Prático':
        return row['AP']
    else:
        return ''


def identifica_modulacao(row):
    if row['Tipo de Atividade'] == 'Teórico':
        return row['MT']
    elif row['Tipo de Atividade'] == 'Prático':
        return row['MP']
    else:
        return ''


def prepara_dados(arquivo, dict_disciplina, lista_professores, df_listas):
    # Lê o arquivo
    dataframe = pandas.read_excel(arquivo, engine='openpyxl')
    dataframe = dataframe[['ESCOLA', 'CURSO', 'DISCIPLINA', 'PERÍODO', 'MATRIZ', 'AT', 'AP', 'MT', 'MP']]

    # Renomeia as colunas
    dataframe.rename(columns={
        'ESCOLA': 'Escola / Campus',
        'CURSO': 'Nome do Curso',
        'DISCIPLINA': 'Nome da disciplina',
        'PERÍODO': 'Período',
        'MATRIZ': 'Matriz',
    }, inplace=True)

    # Filtra apenas as matrizes / períodos que terão oferta
    dataframe_filtrado = pandas.DataFrame()
    ofertas = ['2021-1', '2021-1', '2021-1', '2020-2', '2020-1', '2019-2',
               '2019-1', '2018-2', '2018-1', '2017-2', '2017-1', '2016-2']  # indice + 1 = periodo
    for i, oferta in enumerate(ofertas):
        dataframe_busca = dataframe.loc[(dataframe['Matriz'].str.contains(oferta)) & (dataframe['Período'] == (i + 1))]
        dataframe_filtrado = pandas.concat([dataframe_filtrado, dataframe_busca], axis=0, ignore_index=True)

    dataframe = dataframe_filtrado

    # Insere todas as colunas do modelo
    colunas_para_adicionar = [
        'Cód. do Curso', 'Cód. da Disciplina', 'Turma', 'Turno', 'Presencial',
        'Tipo de Atividade', 'Modulação pela Matriz', 'Previsão do número de alunos',
        'Retorno Presencial?', 'Hy Flex?',
        'Usa Sala Teórica, Metodologia Ativa ou Laboratório?',
        'Qual Ambiente de Aprendizagem? (Laboratório)',
        'Software (Especificar nomes dos softwares)',
        'Possui unificação com outro curso? (Indicar o curso/turma)',
        'Indicar curso base/pagante em caso de unificação.', 'Docente 2022.1',
        'Professor é grupo de risco?',
        'Indicar como eletiva', 'Escola/instituição?', 'Observação', 'Carga Horária'
    ]
    dataframe = \
        pandas.concat([dataframe, pandas.DataFrame(columns=colunas_para_adicionar)], axis=0, ignore_index=True)

    # Aplica o dicionário de disciplinas para preencher os códigos
    dataframe['Cód. da Disciplina'] = dataframe.apply(
        lambda row: dict_disciplina.get(row['Nome da disciplina'], ''), axis=1)

    # Identifica modulação
    dataframe['AT'].astype(float).astype('Int32')
    dataframe['AP'].astype(float).astype('Int32')
    dataframe['Tipo de Atividade'] = dataframe.apply(
        lambda row: identifica_tipo(row['AT'], row['AP']), axis=1)

    # Separa modulação mista em linhas diferentes
    # Duplica todas as disciplinas com modulação 'Ambos', alterando a cópia para 'Teórico'
    dataframe = dataframe.append(
        dataframe.loc[
            dataframe['Tipo de Atividade'] == 'Ambos'
            ].assign(**{'Tipo de Atividade': 'Teórico'}), ignore_index=True)
    # Substitui as disciplinas 'Ambos' originais por 'Prático'
    dataframe['Tipo de Atividade'].replace({'Ambos': 'Prático'}, inplace=True)

    # Identifica a carga horária
    dataframe['Carga Horária'] = dataframe.apply(
        lambda row: identifica_ch(row), axis=1)

    # Identifica a modulação
    dataframe['Modulação pela Matriz'] = dataframe.apply(
        lambda row: identifica_modulacao(row), axis=1)

    # Reordena as colunas
    dataframe = dataframe[
        [
            'Escola / Campus', 'Cód. do Curso', 'Nome do Curso', 'Cód. da Disciplina', 'Nome da disciplina',
            'Período', 'Turma', 'Matriz', 'Turno', 'Presencial', 'Tipo de Atividade', 'Modulação pela Matriz',
            'Carga Horária', 'Previsão do número de alunos', 'Retorno Presencial?', 'Hy Flex?',
            'Usa Sala Teórica, Metodologia Ativa ou Laboratório?', 'Qual Ambiente de Aprendizagem? (Laboratório)',
            'Software (Especificar nomes dos softwares)',
            'Possui unificação com outro curso? (Indicar o curso/turma)',
            'Indicar curso base/pagante em caso de unificação.', 'Docente 2022.1', 'Professor é grupo de risco?',
            'Indicar como eletiva', 'Escola/instituição?', 'Observação',
        ]
    ]

    # Reordena as linhas
    dataframe.sort_values(['Escola / Campus', 'Nome do Curso', 'Período'], inplace=True)

    # Salva um arquivo por curso
    # Seleciona os diferentes valores para Escola + Curso
    lista_cursos = dataframe[['Escola / Campus', 'Nome do Curso']]
    lista_cursos = lista_cursos.drop_duplicates()
    lista_cursos = lista_cursos.values.tolist()

    for escola, curso in lista_cursos:
        # Cria o dataframe filtrado e o arquivo de saída
        dataframe_por_vinculo = dataframe.loc[
            (dataframe['Escola / Campus'] == escola) & (dataframe['Nome do Curso'] == curso)]
        nome_arquivo_saida = f'Planilha Turmas/{escola}/Planilha Turmas - {curso}.xlsx'

        salva_planilha_turmas(dataframe_por_vinculo, nome_arquivo_saida, lista_professores, df_listas)


def main():
    Tk().withdraw()

    # Arquivo para popular as listas de preenchimento (validação do excel)
    arquivo_listas = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o template da planilha turmas')
    df_listas = cria_df_listas(arquivo_listas)

    # Arquivo para popular a lista de professores
    arquivo_professores = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a base de professores ativos')
    lista_professores = cria_lista_professores(arquivo_professores)

    # Arquivo para descobrir o código da disciplina a partir do nome
    arquivo_cod_disciplinas = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o dicionário de códigos de disciplina')
    dict_disciplinas = cria_dicionario_disciplinas(arquivo_cod_disciplinas)

    # Arquivo para receber todos os dados das resoluções do curso
    arquivo_resolucoes = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o arquivo de resoluções')

    print('Processando...')
    prepara_dados(arquivo_resolucoes, dict_disciplinas, lista_professores, df_listas)

    input('Arquivos gerados com sucesso!')


if __name__ == '__main__':
    main()
