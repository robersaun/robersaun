from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas as pd


def encontra_periodo(cod_turma):
    if cod_turma.__contains__('EAD'):
        return 'EAD'
    if not cod_turma.__contains__('-'):
        return ''

    cod_turma = cod_turma.split('-')[1].strip()
    periodo = ''
    for ch in cod_turma:
        if ch in '0123456789':
            periodo += ch
        else:
            break
    return periodo


def encontra_cr(turma):
    if turma.__contains__('EAD'):
        return 'EAD'
    elif not turma.__contains__('-'):
        return ''
    else:
        turno = turma.split('-')[2].strip()
        sigla = turma.split('-')[0].strip()
        validador = sigla + ' ' + turno
        return dict_sigla_cr.get(validador, '')


def encontra_curso(cod_turma):
    cod_turma = cod_turma.split('-')[0].strip()
    curso = dict_sigla_curso.get(cod_turma, '')
    return curso


dict_sigla_cr = {
    'CAUR M': '2244',
    'CAUR N': '2303',
    'CBAV N': '1031087',
    'CCAU M': '1031089',
    'CCAU N': '1031090',
    'CCPP M': '2225',
    'CCPP N': '2232',
    'CPVS M': '2256',
    'CDES M': '2271',
    'CDES N': '2272',
    'CDPD N': '2354',
    'CDDG T': '2342',
    'CJOR M': '2236',
    'CJOR N': '2237',
    'CMUS N': '2215',
    'CMPM N': '2283',
    'CRPU M': '2238',
    'CRPU N': '1031183',
    'CTEA N': '2308',
    'CDIR M': '2022',
    'CDIR N': '2023',
    'CBCS N': '1031080',
    'CFIL M': '2304',
    'CFIL N': '1031182',
    'CBHI M': '1031078',
    'CBHI N': '1031079',
    'CBLI I': '1031081',
    'CTEO M': '2219',
    'CTEO N': '2219',
    'CICH I': '2275',
    'CLPI M': '2203',
    'CLPI N': '2203',
    'CLBI M': '1031181',
    'CLBI N': '2168',
    'CSOC N': '2221',
    'CEDF M': '2206',
    'CEDF N': '2312',
    'CLFL M': '2202',
    'CLFL N': '2302',
    'CLFS M': '2349',
    'CLFS N': '2349',
    'CHIS M': '2214',
    'CHIS N': '2214',
    'CMAT M': '1031465',
    'CMAT N': '2249',
    'CQUI M': '2121',
    'CQUI N': '2121',
    'CPED M': '2201',
    'CPED N': '2210',
    'CSSC M': '2120',
    'CSSC N': '2220',
    'LAD M': '115',
    'LAD N': '2501',
    'LDI M': '2518',
    'LDI N': '2502',
    'LPO N': '119',
    'LES N': '111148',
    'LFS M': '111154',
    'LME I': '2530',
    'LPS M': '121',
    'LPS N': '124',
    'LTO M': '114',
    'GBF M': '114165',
    'GFL M': '2603',
    'CAGR M': '2078',
    'CAGR N': '2079',
    'CBED M': '2212',
    'CBED N': '2311',
    'CIS I': '2295',
    'CBIT M': '2074',
    'CCBI M': '2360',
    'CCBI T': '2160',
    'CEFB M': '1031582',
    'CEFB N': '1031583',
    'CENF M': '2069',
    'CFAR M': '2167',
    'CFIS M': '2066',
    'CMED I': '2062',
    'CMVT M': '2076',
    'CNUT M': '2070',
    'CODO I': '2064',
    'CPSI M': '2063',
    'CPSI N': '2071',
    'CGAS M': '2072',
    'CGAS N': '1060',
    'CADM M': '2028',
    'CADM N': '2029',
    'CADI M': '2239',
    'CBND M': '1031461',
    'CBND N': '1031467',
    'CBNI M': '1031462',
    'CBNI N': '1031468',
    'CCOT M': '2229',
    'CCOT N': '2230',
    'CECO M': '2036',
    'CECO N': '2035',
    'CIBP I': '2293',
    'CMKT M': '2129',
    'CMKT N': '2128',
    'CTUR N': '2227',
    'CBJD M': '1031092',
    'CCBS M': '1031469',
    'CCBS N': '2287',
    'CCCO M': '2245',
    'CCCO N': '2045',
    'CEAM M': '2246',
    'CEBI M': '2222',
    'CEBI N': '2223',
    'CECV M': '2243',
    'CECV N': '2343',
    'CECP M': '2248',
    'CECP N': '2348',
    'CECA M': '2158',
    'CECA N': '2159',
    'CEPD M': '2241',
    'CEPD N': '2341',
    'CESF M': '2263',
    'CEEE M': '2355',
    'CEEE N': '2359',
    'CEET N': '2351',
    'CEEL M': '2140',
    'CEMC M': '2052',
    'CEMC N': '2352',
    'CEMT M': '2269',
    'CEMT N': '2259',
    'CQUE M': '2353',
    'CQUE N': '2253',
    'CBSI M': '2150',
    'CBSI N': '2250',
    'CJDI M': '2145',
    'CJDI N': '2145',
    'CTSI N': '2289',
    'CTAI N': '1031459',
    'TAD N': '2402',
    'TAG M': '2420',
    'TAG N': '2424',
    'TCC N': '2429',
    'TDI M': '2430',
    'TDI N': '2431',
    'TPO N': '2422',
    'TGI N': '112178',
    'TMV M': '2410',
    'TMV N': '112152',
    'TPS M': '112153',
    'TPS N': '2423',
    'CBLI M': '1031081'
}

dict_sigla_curso = {
    'CAUR': 'Arquitetura e Urbanismo',
    'CCAU': 'Cinema e Audiovisual',
    'CCPP': 'Comunicação Social - Hab. Publicidade e Propaganda',
    'CPVS': 'Desenho Industrial - Hab.: Programação Visual',
    'CDES': 'Design',
    'CJOR': 'Jornalismo',
    'CRPU': 'Relações Públicas',
    'CBAV': 'Bacharelado em Artes Visuais',
    'CDPD': 'Design de Produto',
    'CMUS': 'Licenciatura em Música',
    'CMPM': 'Produção Musical',
    'CTEA': 'Teatro',
    'CDDG': 'Design Digital',
    'CDIR': 'Direito',
    'CBLI': 'Bacharelado em Letras Inglês Internacional',
    'CICH': 'Bacharelado Interdisciplinar em Ciências e Humanidades',
    'CFIL': 'Bacharelado em Filosofia',
    'CBHI': 'Bacharelado em História ',
    'CTEO': 'Bacharelado em Teologia',
    'CLPI': 'Letras-Português-Inglês',
    'CLBI': 'Licenciatura em Ciências Biológicas',
    'CEDF': 'Licenciatura em Educação Física',
    'CLFL': 'Licenciatura em Filosofia',
    'CLFS': 'Licenciatura em Física',
    'CHIS': 'Licenciatura em História',
    'CMAT': 'Licenciatura em Matemática',
    'CQUI': 'Licenciatura em Química',
    'CPED': 'Pedagogia',
    'CSSC': 'Serviço Social',
    'CBCS': 'Bacharelado em Ciências Sociais',
    'CSOC': 'Licenciatura em Ciências Sociais',
    'LME': 'Medicina',
    'LAD': 'Administração',
    'LDI': 'Direito',
    'LFS': 'Fisioterapia',
    'LPS': 'Psicologia',
    'LTO': 'Teologia',
    'LPO': 'Engenharia de Produção',
    'LES': 'Engenharia de Software',
    'GBF': 'Bacharelado em Filosofia',
    'GFL': 'Filosofia',
    'CIS': 'Bacharelado Interdisciplinar em Saúde',
    'CMED': 'Medicina',
    'CODO': 'Odontologia',
    'CAGR': 'Agronomia',
    'CBED': 'Bacharelado em Educação Física',
    'CBIT': 'Biotecnologia',
    'CCBI': 'Ciências Biológicas - Bacharelado',
    'CEFB': 'Educação Física',
    'CENF': 'Enfermagem',
    'CFAR': 'Farmácia',
    'CFIS': 'Fisioterapia',
    'CMVT': 'Medicina Veterinária',
    'CNUT': 'Nutrição',
    'CPSI': 'Psicologia',
    'CGAS': 'Superior de Tecnologia em Gastronomia',
    'CIBP': 'International Business Program IBP',
    'CADM': 'Administração',
    'CADI': 'Administração - Internacional',
    'CBND': 'Bacharelado em Negócios Digitais',
    'CBNI': 'Bacharelado em Negócios Internacionais',
    'CCOT': 'Ciências Contábeis',
    'CECO': 'Ciências Econômicas',
    'CMKT': 'Marketing',
    'CTUR': 'Turismo',
    'CBJD': 'Bacharelado em Jogos Digitais',
    'CCBS': 'Cibersegurança',
    'CCCO': 'Ciência da Computação',
    'CEAM': 'Engenharia Ambiental',
    'CEBI': 'Engenharia Biomédica',
    'CECV': 'Engenharia Civil',
    'CECP': 'Engenharia de Computação',
    'CECA': 'Engenharia de Controle e Automação',
    'CEPD': 'Engenharia de Produção',
    'CESF': 'Engenharia de Software',
    'CEEE': 'Engenharia Elétrica - Eixos: Telecomunicações Eletrônica ou Sistema de Potência e Energia',
    'CEEL': 'Engenharia Eletrônica',
    'CEMC': 'Engenharia Mecânica',
    'CEMT': 'Engenharia Mecatrônica',
    'CQUE': 'Engenharia Química',
    'CBSI': 'Sistemas de Informação',
    'CJDI': 'Superior de Tecnologia em Jogos Digitais',
    'CEET': 'Engenharia Elétrica (Ênfase em Telecomunicações)',
    'CTSI': 'Superior de Tecnologia em Segurança da Informação',
    'CTAI': 'Tecnólogo em Automação Industrial',
    'TAG': 'Agronomia',
    'TDI': 'Direito',
    'TMV': 'Medicina Veterinária',
    'TPS': 'Psicologia',
    'TAD': 'Administração',
    'TCC': 'Ciências Contábeis',
    'TPO': 'Engenharia de Produção',
    'TGI': 'Gestão Integrada de Agronegócios',
}


def trata_relatorio_ch(arquivo):
    # -- CH na Base Ouro --
    # Adiciona ch relogio oficial na base ouro
    # professor - carga horária por disciplina.xlsx
    # validador base ouro = cr disciplina + turma + disciplina
    # validador = cr + turma + disciplina

    df = pd.read_excel(arquivo)
    df = df[['CR Curso', 'Turma', 'Disciplina', 'C.H. Re. Of']]
    df['CR Curso'] = pd.to_numeric(df['CR Curso'], errors='coerce')

    df['CR Curso'] = df['CR Curso'].fillna(0.0).astype(int)

    df.insert(0, 'Validador', '', True)
    df['Validador'] = df.apply(lambda row: f"{str(row['CR Curso'])}    {row['Turma']}    {row['Disciplina']}",
                               axis=1)

    df = df[['Validador', 'C.H. Re. Of']]

    return df


def main():
    print('Selecione a Relação de Alunos/Pais Exportação')

    Tk().withdraw()
    relacao_alunos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a Relação de Alunos/Pais Exportação')
    print(f'    {relacao_alunos}')

    print('Selecione o relatório de Alunos Matriculados por Disciplina')

    relatorio_disciplinas = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o relatório de Alunos Matriculados por Disciplina')
    print(f'    {relatorio_disciplinas}')

    print('Selecione o relatório de Professor - Carga Horária por Turma e Disciplina')

    relatorio_ch = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o relatório de Professor - Carga Horária por Turma e Disciplina')
    print(f'    {relatorio_ch}')
    print('Gerando a base ouro...')

    df_alunos = pd.read_excel(relacao_alunos)
    df_alunos = df_alunos[[
        'Estabelecimento', 'Escola', 'Centro de Resultado', 'Curso', 'Série', 'Matrícula',
        'Nome Completo', 'CPF', 'Data de Nascimento', 'E-mail', 'Telefone Celular',
        'Situação Acadêmica', 'Tipo de Ingresso', 'Turma', 'Turno', 'Gênero'
    ]]
    df_alunos = df_alunos.loc[df_alunos["Situação Acadêmica"] == 'Matriculado Curso Normal']
    df_alunos = df_alunos.drop_duplicates()

    df_disciplinas = pd.read_excel(relatorio_disciplinas)
    df_disciplinas = df_disciplinas[[
        'CODIGO', 'DT_CADASTRO_CONTRATO', 'TURMA_BASE', 'DISCIPLINA', 'TURMA_DISCIPLINA', 'DIVISAO', 'DIVISAO2'
    ]]
    df_disciplinas = df_disciplinas.drop_duplicates()

    df_ch = trata_relatorio_ch(relatorio_ch)

    print('Juntando dados...')

    df_joined = pd.merge(
        left=df_alunos, right=df_disciplinas, left_on=['Matrícula', 'Turma'], right_on=['CODIGO', 'TURMA_BASE']
    )

    # modificando o dataframe
    df_joined = df_joined[['Estabelecimento', 'Escola', 'Centro de Resultado', 'Curso', 'Série', 'Matrícula',
                           'Nome Completo', 'CPF', 'Data de Nascimento', 'E-mail', 'Telefone Celular',
                           'Situação Acadêmica', 'Tipo de Ingresso', 'Turma', 'Turno', 'Gênero',
                           # dados disciplina
                           'DT_CADASTRO_CONTRATO', 'DISCIPLINA', 'TURMA_DISCIPLINA', 'DIVISAO', 'DIVISAO2']]

    df_joined.rename(columns={
        'Série': 'Período Aluno',
        'Turma': 'Turma Aluno',
        'Centro de Resultado': 'CR Aluno',
        'Curso': 'Curso Aluno',
        'Disciplina': 'DISCIPLINA',
        'Turma Destino': 'TURMA_DISCIPLINA',
    }, inplace=True)

    print('Calculando...')

    df_joined.insert(17, 'Curso_Disciplina', '', True)
    df_joined['Curso_Disciplina'] = df_joined.apply(lambda row: encontra_curso(row['TURMA_DISCIPLINA']), axis=1)

    df_joined.insert(18, 'Período_Disciplina', '', True)
    df_joined['Período_Disciplina'] = df_joined.apply(lambda row: encontra_periodo(row['TURMA_DISCIPLINA']), axis=1)

    df_joined.insert(19, 'CR_Disciplina', '', True)
    df_joined['CR_Disciplina'] = df_joined.apply(
        lambda row: encontra_cr(row['TURMA_DISCIPLINA']),
        axis=1)

    # Remove o início do nome do estabelecimento
    df_joined['Estabelecimento'] = df_joined['Estabelecimento'].str.replace(
        'Pontifícia Universidade Católica do Paraná - ', '')

    # Remove o início do nome da escola
    df_joined['Escola'] = df_joined['Escola'] \
        .str.replace('Escola de ', '') \
        .str.replace('Escola ', '')

    # Corrige Belas Artes
    df_joined['Escola'] = df_joined['Escola'] \
        .str.replace('Comunicação e Artes', 'Belas Artes') \
        .str.replace('Arquitetura e Design', 'Belas Artes')

    # Altera a escola para o nome do campus fora de sede
    df_joined.loc[
        (df_joined['Estabelecimento'] == 'Londrina') |
        (df_joined['Estabelecimento'] == 'Maringá') |
        (df_joined['Estabelecimento'] == 'Toledo'),
        'Escola'] = df_joined['Estabelecimento']

    # No período do aluno deixar só os números
    df_joined['Período Aluno'] = df_joined['Período Aluno'] \
        .str.replace('º Periodo', '') \
        .str.replace('º Período', '')

    # Preenche a coluna Gênero com "Não informado" quando estiver vazia
    df_joined["Gênero"].fillna("Não informado", inplace=True)

    # Remove espaços antes e depois dos nomes
    df_joined['Nome Completo'] = df_joined['Nome Completo'].str.strip()

    # Insere validador para a inclusão da coluna de carga horária
    df_joined.insert(0, 'Validador', '', True)
    df_joined['Validador'] = df_joined.apply(
        lambda row: f"{str(row['CR_Disciplina'])}    {row['TURMA_DISCIPLINA']}    {row['DISCIPLINA']}", axis=1)

    df_joined = pd.merge(left=df_joined, right=df_ch, left_on='Validador', right_on='Validador', how='left')
    del df_joined['Validador']

    # Remove linhas duplicadas
    df_joined = df_joined.drop_duplicates()

    print('Criando arquivos de saída...')

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('base_ouro_completa.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df_joined.to_excel(writer, sheet_name='Sheet1', index=False)

    print('Salvando arquivos...')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    print('Geração de arquivos finalizada!')


if __name__ == '__main__':
    main()
