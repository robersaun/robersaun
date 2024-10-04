"""Base Ouro (Versão 2)

Script para gerar a Base Ouro, um relatório do vínculo de todos os alunos e disciplinas.
Recebe como entrada dois relatórios extraídos do sistema Prime
e cria um arquivo único em excel.

Não é mais utilizado, versão descontinuada.
"""
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


def encontra_cr(curso, turno):
    if turno.__contains__('EAD'):
        return 'EAD'
    elif not turno.__contains__('-'):
        return ''
    else:
        turno = turno.split('-')[2].strip()
        curso = curso + ' ' + turno
        return dict_curso_cr.get(curso, '')


def encontra_curso(cod_turma):
    cod_turma = cod_turma.split('-')[0].strip()
    curso = dict_turma_curso.get(cod_turma, '')
    return curso


dict_curso_cr = {
    'Administração - Internacional M': '2239',
    'Administração M': '2028',
    'Administração N': '2029',
    'Agronomia M': '2078',
    'Agronomia N': '2079',
    'Arquitetura e Urbanismo M': '2244',
    'Arquitetura e Urbanismo N': '2303',
    'Arquitetura e Urbanismo T': '2255',
    'Bacharelado em Educação Física M': '2212',
    'Bacharelado em Educação Física N': '2311',
    'Biotecnologia M': '2074',
    'Sistemas de Informação N': '2250',
    'Ciências Biológicas - Bacharelado T': '2160',
    'Ciências Biológicas - Bacharelado M': '2360',
    'Cibersegurança N': '2287',
    'Ciência da Computação M': '2245',
    'Ciência da Computação N': '2345',
    'Comunicação Social - Hab.: Publicidade e Propaganda M': '2225',
    'Comunicação Social - Hab.: Publicidade e Propaganda N': '2232',
    'Ciências Contábeis N': '2230',
    'Ciências Contábeis M': '2229',
    'Design Digital T': '2342',
    'Design M': '2271',
    'Design N': '2272',
    'Design Gráfico M': '2356',
    'Direito M': '2022',
    'Direito N': '2023',
    'Design de Moda M': '2357',
    'Engenharia de Alimentos M': '2147',
    'Engenharia Ambiental N': '2346',
    'Engenharia Ambiental M': '2246',
    'Engenharia Biomédica M': '2222',
    'Engenharia de Controle e Automação M': '2158',
    'Engenharia de Controle e Automação N': '2159',
    'Ciências Econômicas M': '2036',
    'Ciências Econômicas N': '2035',
    'Engenharia de Computação M': '2248',
    'Engenharia de Computação N': '2348',
    'Engenharia Civil N': '2343',
    'Engenharia Civil M': '2243',
    'Licenciatura em Educação Física M': '2206',
    'Licenciatura em Educação Física N': '2312',
    'Engenharia Elétrica - Eixos: Telecomunicações, Eletrônica ou Sistema de Potência e Energia M': '2355',
    'Engenharia Elétrica - Eixos: Telecomunicações, Eletrônica ou Sistema de Potência e Energia N': '2359',
    'Engenharia Eletrônica M': '2140',
    'Engenharia Elétrica (Ênfase em Telecomunicações) N': '2351',
    'Engenharia Mecânica M': '2252',
    'Engenharia Mecânica N': '2352',
    'Engenharia Mecatrônica M': '2269',
    'Engenharia Mecatrônica N': '2270',
    'Enfermagem M': '2069',
    'Engenharia de Produção N': '2341',
    'Engenharia de Produção M': '2241',
    'Engenharia de Software M': '2263',
    'Farmácia M': '2167',
    'Bacharelado em Filosofia M': '2304',
    'Fisioterapia M': '2066',
    'Superior de Tecnologia em Gastronomia M': '2072',
    'Superior de Tecnologia em Gastronomia N': '1060',
    'Licenciatura em História M': '2314',
    'Licenciatura em História N': '2214',
    'Bacharelado Interdisciplinar em Ciências e Humanidades M': '2275',
    'Bacharelado Interdisciplinar em Saúde I': '2295',
    'Superior de Tecnologia em Jogos Digitais N': '2145',
    'Superior de Tecnologia em Jogos Digitais M': '2344',
    'Jornalismo M': '2236',
    'Jornalismo N': '2237',
    'Licenciatura em Ciências Biológicas N': '2268',
    'Licenciatura em Filosofia N': '2302',
    'Licenciatura em Física N': '2349',
    'Letras-Português-Inglês N': '2203',
    'Letras-Português-Inglês M': '2301',
    'Licenciatura em Matemática N': '2249',
    'Medicina I': '2062',
    'Marketing N': '2128',
    'Marketing M': '2129',
    'Música - Produção Musical N': '2283',
    'Licenciatura em Música N': '2215',
    'Medicina Veterinária M': '2076',
    'Nutrição M': '2070',
    'Odontologia M': '2064',
    'Odontologia I': '2064',
    'Pedagogia M': '2201',
    'Pedagogia N': '2210',
    'Psicologia M': '2063',
    'Psicologia N': '2071',
    'Engenharia Química M': '2353',
    'Engenharia Química N': '2253',
    'Licenciatura em Química N': '2350',
    'Licenciatura em Química M': '2121',
    'Relações Públicas M': '2238',
    'Ciências Sociais N': '2221',
    'Serviço Social N': '2220',
    'Teatro N': '2308',
    'Bacharelado em Teologia M': '2319',
    'Bacharelado em Teologia N': '2219'
}
dict_turma_curso = {
    'CADI': 'Administração - Internacional',
    'CADM': 'Administração',
    'CAGR': 'Agronomia',
    'CAUR': 'Arquitetura e Urbanismo',
    'CBED': 'Bacharelado em Educação Física',
    'CBIT': 'Biotecnologia',
    'CBSI': 'Sistemas de Informação',
    'CCBI': 'Ciências Biológicas - Bacharelado',
    'CCBS': 'Cibersegurança',
    'CCCO': 'Ciência da Computação',
    'CCCP': 'Comunicação Social - Hab.: Publicidade e Propaganda',
    'CCOT': 'Ciências Contábeis',
    'CCPP': 'Comunicação Social - Hab.: Publicidade e Propaganda',
    'CDDG': 'Design Digital',
    'CDES': 'Design',
    'CDGR': 'Design Gráfico',
    'CDIR': 'Direito',
    'CDMD': 'Design de Moda',
    'CEAL': 'Engenharia de Alimentos',
    'CEAM': 'Engenharia Ambiental',
    'CEBI': 'Engenharia Biomédica',
    'CECA': 'Engenharia de Controle e Automação',
    'CECO': 'Ciências Econômicas',
    'CECP': 'Engenharia de Computação',
    'CECV': 'Engenharia Civil',
    'CEDF': 'Licenciatura em Educação Física',
    'CEEE': 'Engenharia Elétrica - Eixos: Telecomunicações, Eletrônica ou Sistema de Potência e Energia',
    'CEEL': 'Engenharia Eletrônica',
    'CEET': 'Engenharia Elétrica (Ênfase em Telecomunicações)',
    'CEMC': 'Engenharia Mecânica',
    'CEMT': 'Engenharia Mecatrônica',
    'CENF': 'Enfermagem',
    'CEPD': 'Engenharia de Produção',
    'CESF': 'Engenharia de Software',
    'CFAR': 'Farmácia',
    'CFIL': 'Bacharelado em Filosofia',
    'CFIS': 'Fisioterapia',
    'CGAS': 'Superior de Tecnologia em Gastronomia',
    'CHIS': 'Licenciatura em História',
    'CICH': 'Bacharelado Interdisciplinar em Ciências e Humanidades',
    'CIS': 'Bacharelado Interdisciplinar em Saúde',
    'CJDI': 'Superior de Tecnologia em Jogos Digitais',
    'CJOR': 'Jornalismo',
    'CLBI': 'Licenciatura em Ciências Biológicas',
    'CLFL': 'Licenciatura em Filosofia',
    'CLFS': 'Licenciatura em Física',
    'CLPI': 'Letras-Português-Inglês',
    'CMAT': 'Licenciatura em Matemática',
    'CMED': 'Medicina',
    'CMKT': 'Marketing',
    'CMPM': 'Música - Produção Musical',
    'CMUS': 'Licenciatura em Música',
    'CMVT': 'Medicina Veterinária',
    'CNUT': 'Nutrição',
    'CODO': 'Odontologia',
    'CPED': 'Pedagogia',
    'CPSI': 'Psicologia',
    'CQUE': 'Engenharia Química',
    'CQUI': 'Licenciatura em Química',
    'CRPU': 'Relações Públicas',
    'CSOC': 'Ciências Sociais',
    'CSSC': 'Serviço Social',
    'CTEA': 'Teatro',
    'CTEO': 'Bacharelado em Teologia',
    'ENGENHARIA AMPERE': 'Engenharia',
    'ENGENHARIA ARISTOTELES': 'Engenharia',
    'ENGENHARIA ARQUIMEDES': 'Engenharia',
    'ENGENHARIA AZIMOV': 'Engenharia',
    'ENGENHARIA BERNOULLI KEPLER': 'Engenharia',
    'ENGENHARIA BOHR': 'Engenharia',
    'ENGENHARIA BOYLE': 'Engenharia',
    'ENGENHARIA DARWIN': 'Engenharia',
    'ENGENHARIA FLEMING': 'Engenharia',
    'ENGENHARIA GALILEU FARADAY': 'Engenharia',
    'ENGENHARIA GATES': 'Engenharia',
    'ENGENHARIA GAUSS PLANCK': 'Engenharia',
    'ENGENHARIA HAWKING': 'Engenharia',
    'ENGENHARIA HEISENBERG': 'Engenharia',
    'ENGENHARIA HERTZ NOBEL': 'Engenharia',
    'ENGENHARIA HIGGS': 'Engenharia',
    'ENGENHARIA JOULE': 'Engenharia',
    'ENGENHARIA LATTES': 'Engenharia',
    'ENGENHARIA LEIBNITZ': 'Engenharia',
    'ENGENHARIA MARQUES': 'Engenharia',
    'ENGENHARIA MAXWELL': 'Engenharia',
    'ENGENHARIA MENDEL THOMPSON': 'Engenharia',
    'ENGENHARIA MORSE': 'Engenharia',
    'ENGENHARIA PAULING': 'Engenharia',
    'LEA': 'Letras-Português-Inglês',
    'LETTC': 'Letras-Português-Inglês',
    'GPS': 'Psicologia',
    'LDI': 'Direito',
    'TDI': 'Direito',
    'TAG': 'Agronomia',
    'TCC': 'Ciências Contábeis',
    'LAD': 'Administração',
    'LME': 'Medicina',
    'LPS': 'Psicologia',
    'GDI': 'Direito',
    'LPO': 'Engenharia de Produção',
    'TMV': 'Medicina Veterinária',
    'TPS': 'Psicologia',
    'GFL': 'Filosofia',
    'LTO': 'Teologia',
    'TPO': 'Engenharia de Produção',
    'TAD': 'Administração',
    'LCC': 'Ciências Contábeis',
    'CDPD': 'Design de Produto',
    'GAD': 'Administração',
    'LES': 'Engenharia de Software',
    'CECA 1U': 'Engenharia de Computação',
    'CIBP': 'International Business Program IBP',
    'LFS': 'Fisioterapia',
    'TBB': 'Ciências Biológicas',
    'CPVS': 'Desenho Industrial - Hab.: Programação Visual',
    'CCJR': 'Comunicação Social - Hab.: Jornalismo',
    'CEFL': 'Engenharia Florestal',
    'CTUR': 'Turismo',
    'CTSI': 'Superior de Tecnologia em Segurança da Informação',
}


def main():
    print('Selecione a Relação de Alunos/Pais Exportação')

    Tk().withdraw()
    relacao_alunos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a Relação de Alunos/Pais Exportação')

    print('Selecione o relatório de Alunos Matriculados por Disciplina')

    relatorio_disciplinas = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o relatório de Alunos Matriculados por Disciplina')

    df_alunos = pd.read_excel(relacao_alunos)
    df_alunos = df_alunos[[
        'Estabelecimento', 'Escola', 'Centro de Resultado', 'Curso', 'Série', 'Matrícula',
        'Nome Completo', 'CPF', 'Data de Nascimento', 'E-mail', 'Telefone Celular',
        'Situação Acadêmica', 'Tipo de Ingresso', 'Turma', 'Turno'
    ]]
    df_alunos = df_alunos[df_alunos['Situação Acadêmica'] == 'Matriculado Curso Normal']
    df_alunos.drop_duplicates()

    df_disciplinas = pd.read_excel(relatorio_disciplinas)
    df_disciplinas = df_disciplinas[[
        'Código', 'Disciplina', 'Turma Destino'
    ]]
    df_disciplinas.drop_duplicates()

    print('Juntando dados...')

    df_joined = pd.merge(left=df_alunos, right=df_disciplinas, left_on='Matrícula', right_on='Código')

    # modificando o dataframe
    df_joined = df_joined[['Estabelecimento', 'Escola', 'Centro de Resultado', 'Curso', 'Série', 'Matrícula',
                           'Nome Completo', 'CPF', 'Data de Nascimento', 'E-mail', 'Telefone Celular',
                           'Situação Acadêmica',
                           'Tipo de Ingresso', 'Turma', 'Turno',
                           # dados disciplina
                           'Disciplina', 'Turma Destino']]

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
        lambda row: encontra_cr(row['Curso_Disciplina'], row['TURMA_DISCIPLINA']),
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

    df_toledo = df_joined[df_joined['Estabelecimento'].str.contains('Toledo')]
    df_maringa = df_joined[df_joined['Estabelecimento'].str.contains('Maringá')]
    df_londrina = df_joined[df_joined['Estabelecimento'].str.contains('Londrina')]
    df_curitiba = df_joined[df_joined['Estabelecimento'].str.contains('Curitiba')]

    print('Criando arquivos de saída...')

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('base_ouro_completa.xlsx', engine='xlsxwriter')
    writer_t = pd.ExcelWriter('base_ouro_toledo.xlsx', engine='xlsxwriter')
    writer_m = pd.ExcelWriter('base_ouro_maringa.xlsx', engine='xlsxwriter')
    writer_l = pd.ExcelWriter('base_ouro_londrina.xlsx', engine='xlsxwriter')
    writer_c = pd.ExcelWriter('base_ouro_curitiba.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df_joined.to_excel(writer, sheet_name='Sheet1', index=False)
    df_toledo.to_excel(writer_t, sheet_name='Sheet1', index=False)
    df_maringa.to_excel(writer_m, sheet_name='Sheet1', index=False)
    df_londrina.to_excel(writer_l, sheet_name='Sheet1', index=False)
    df_curitiba.to_excel(writer_c, sheet_name='Sheet1', index=False)

    print('Salvando arquivos...')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    writer_t.save()
    writer_m.save()
    writer_l.save()
    writer_c.save()

    print('Geração de arquivos finalizada!')


if __name__ == '__main__':
    main()
