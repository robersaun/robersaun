'''
Script realiza uma verificação no arquivo das matrizes, buscando incosistencia de crédito, disciplinas não encontras e erros no cod no curso

Desenvolvido por Vinicius Tozo
Última atualização: 13/08/2021
'''

import pandas

relacao_cursos = {'2275': 'CICH', '2244': 'CAUR', '2303': 'CAUR', '2255': 'CAUR', '2231': 'CCJR', '2225': 'CCPP',
                  '2232': 'CCPP', '2226': 'CCRP', '2256': 'CPVS', '2271': 'CDES', '2272': 'CDES', '2357': 'CDMD',
                  '2354': 'CDPD', '2342': 'CDDG', '2356': 'CDGR', '2236': 'CJOR', '2224': 'CCJR', '2215': 'CMUS',
                  '2283': 'CMPM', '2238': 'CRPU', '2308': 'CTEA', '2078': 'CAGR', '2079': 'CAGR', '2074': 'CBIT',
                  '2360': 'CCBI', '2072': 'CGAS', '1060': 'CGAS', '2212': 'CBED', '2311': 'CBED', '2069': 'CENF',
                  '2369': 'CENF', '2167': 'CFAR', '2066': 'CFIS', '2166': 'CFIS', '2076': 'CMVT', '2070': 'CNUT',
                  '2064': 'CODO', '2063': 'CPSI', '2071': 'CPSI', '2022': 'CDIR', '2023': 'CDIR', '2268': 'CLBI',
                  '2221': 'CSOC', '2206': 'CEDF', '2312': 'CEDF', '2202': 'CLFL', '2302': 'CLFL', '2304': 'CFIL',
                  '2349': 'CLFS', '2314': 'CHIS', '2214': 'CHIS', '2301': 'CLPI', '2203': 'CLPI', '2249': 'CMAT',
                  '2201': 'CPED', '2210': 'CPED', '2121': 'CQUI', '2350': 'CQUI', '2220': 'CSSC', '2319': 'CTEO',
                  '2219': 'CTEO', '115': 'LAD', '2501': 'LAD', '111': 'LCC', '2502': 'LDI', '2518': 'LDI', '119': 'LPO',
                  '111148': 'LES', '2530': 'LME', '121': 'LPS', '124': 'LPS', '114': 'LTO', '2610': 'GAD',
                  '2611': 'GAD', '2605': 'GDI', '2606': 'GDI', '2603': 'GFL', '2614': 'GPS', '2615': 'GPS',
                  '2295': 'CIS', '2062': 'CMED', '2028': 'CADM', '2029': 'CADM', '2239': 'CADI', '2229': 'CCOT',
                  '2230': 'CCOT', '2036': 'CECO', '2035': 'CECO', '2293': 'CIBP', '2129': 'CMKT', '2128': 'CMKT',
                  '2227': 'CTUR', '2287': 'CCBS', '2245': 'CCCO', '2345': 'CCCO', '2246': 'CEAM', '2346': 'CEAM',
                  '2222': 'CEBI', '2223': 'CEBI', '2243': 'CECV', '2343': 'CECV', '2248': 'CECP', '2348': 'CECP',
                  '2147': 'CEAL', '2158': 'CECA', '2159': 'CECA', '2261': 'CEMT', '2241': 'CEPD', '2341': 'CEPD',
                  '2263': 'CESF', '2351': 'CEET', '2140': 'CEEL', '2252': 'CEMC', '2352': 'CEMC', '2353': 'CQUE',
                  '2253': 'CQUE', '2344': 'CJDI', '2145': 'CJDI', '2250': 'CBSI', '2402': 'TAD', '2420': 'TAG',
                  '2424': 'TAG', '2404': 'TBB', '2429': 'TCC', '2430': 'TDI', '2431': 'TDI', '2422': 'TPO',
                  '2410': 'TMV', '112152': 'TMV', '112153': 'TPS', '2423': 'TPS', '2289': 'CTSI', '2355': 'CEEE',
                  '2359': 'CEEE', '2277': 'CPED', '2269': 'CEMT', '2270': 'CEMT', '2160': 'CCBI', '2075': 'CEFL',
                  # Corrigidos depois, não aparecem na base ouro
                  '2103': 'CAUR', '2242': 'CDES', '2247': 'CDMD', '2254': 'CDPD', '2060': 'CCBI', '2093': 'CMVT',
                  '2094': 'CBIT', '2095': 'CEFL', '2169': 'CENF', '2193': 'CMVT', '2291': 'CAGR', '2391': 'CAGR',
                  '2421': 'CENF', '2425': 'CFAR', '2601': 'CNUT', '2233': 'CCRP', '2084': 'CDIR', '2085': 'CDIR',
                  '2204': 'Extinto', '2205': 'CTEO', '2211': 'Extinto', '2216': 'Extinto', '2081': 'CADM',
                  '2130': 'Extinto', '2131': 'Extinto', '2132': 'Extinto', '2133': 'Extinto', '2134': 'Extinto',
                  '2135': 'Extinto', '2207': 'Extinto', '2288': 'CCOT', '117': 'CEPD', '123': 'Extinto',
                  '2048': 'CECP', '2052': 'CEMC', '2058': 'CEMT', '2143': 'CECV', '2144': 'Extinto', '2151': 'Extinto',
                  '2240': 'CEAL', '2258': 'CEMT', '2358': 'CEMT', '2426': 'CEAM', '2110': 'Extinto', '2111': 'Extinto',
                  '2368': 'CCBI', '2120': 'CSSC', '2347': 'CDMD', '2503': 'CBSI', '2531': 'Extinto', '2532': 'Extinto',
                  '2082': 'CMVT', '2237': 'CJOR', '2228': 'CTUR', '2276': 'CEFL', '2273': 'Extinto', '2274': 'Extinto',
                  '2200': 'CTSI', '2257': 'CTSI', '2260': 'CBSI', '2262': 'CEMT', '2264': 'CESF', '2278': 'CCOT',
                  '2279': 'CECO', '2280': 'CMKT', '30000': 'Intercâmbio'
                  }


def trata_disciplina_relatorio(codigo, nome):
    # Tratamento de exceções
    if nome == "Fundamentos de Sistemas Ciber-Físicos":
        nome = "Fundamentos de Sistemas Ciberfísicos"

    return str(codigo) + " - " + nome


def main():
    # Arquivo de saída
    out_txt = open("out.txt", "w", encoding="utf-8")
    print(";".join([
        'Tipo de erro', 'CR', 'Curso', 'Disciplina', 'Abreviacao', 'Ano',
        'Créditos no relatório', 'Matriz', 'Créditos na matriz', 'Classificação', 'Turma'
    ]), file=out_txt)

    pandas.set_option('display.max_columns', None)

    # Relatório gerado por semestre
    caminho_relatorio = "relatorio_Relatorio ofertas 2011 a 2020_1.xlsx"
    relatorio = pandas.read_excel(caminho_relatorio)
    relatorio.columns = [
        'Escola', 'Cod Curso', 'Nome Curso', 'Ano da oferta', 'Sem da oferta',
        'Periodo', 'Turma', 'Cod Disciplina BD', 'Cod Disciplina',
        'Nome Disciplina', 'Mod. Teorica', 'Mod. Pratica', 'Aulas teoricas',
        'Aulas praticas', 'Creditos', 'Total HA', 'Total HR',
        'Tipo de fechamento', 'Campus', 'Vagas presenciais',
        'Vagas nao presenciais', 'Tem professor', 'Tem Ensalamento',
        'Modulacao Ensalada', 'Divisao Turma'
    ]
    relatorio = relatorio[[
        'Cod Curso', 'Nome Curso', 'Ano da oferta', 'Sem da oferta',
        'Cod Disciplina', 'Nome Disciplina', 'Creditos', 'Turma'
    ]]

    # Arquivo gerado pelo script "cria-excel-unificado.py" contendo todas as matrizes da pasta DAC
    excel_unificado = "excel-unificado.xlsx"
    unificado = pandas.read_excel(excel_unificado)
    unificado.columns = ['Disciplina', 'Classificação', 'Créditos', 'Grupo', 'CH Teórica', 'CH Prática', 'CH Oficial',
                         'CH Relógio Oficial', 'Curso', 'Matriz', 'Abreviação', 'Ano', 'Semestre']

    encontrados = 0
    nao_encontrados = 0
    creditos_inconsistentes = 0
    total = len(relatorio.index)
    cursos_sem_abreviacao = []
    for index, linha_relatorio in relatorio.iterrows():

        # Pega as informações da linha atual
        nome_disciplina = linha_relatorio["Nome Disciplina"]
        disciplina_relatorio = trata_disciplina_relatorio(linha_relatorio['Cod Disciplina'], nome_disciplina)
        curso_relatorio = str(linha_relatorio['Cod Curso'])
        abreviacao_relatorio = relacao_cursos.get(curso_relatorio, '')
        ano_relatorio = str(linha_relatorio["Ano da oferta"])

        # Se não encontrar o CR
        if abreviacao_relatorio == "" and curso_relatorio not in cursos_sem_abreviacao:
            print(";".join([
                'Abreviação não encontrada para o CR informado', curso_relatorio, linha_relatorio["Nome Curso"],
                nome_disciplina,
                abreviacao_relatorio, ano_relatorio, '-', '-', '-', '-', str(linha_relatorio['Turma'])
            ]), file=out_txt)
            cursos_sem_abreviacao.append(curso_relatorio)
            continue

        # Busca por cod_disciplina + nome_disciplina por abreviação de curso
        busca_unificado = unificado.loc[(unificado['Disciplina'] == disciplina_relatorio) &
                                        (abreviacao_relatorio == unificado['Abreviação'])]

        # Se não encontrar
        if busca_unificado.empty:
            nao_encontrados += 1
            print(";".join([
                'Disciplina não encontrada', str(linha_relatorio['Cod Curso']), linha_relatorio['Nome Curso'],
                disciplina_relatorio, abreviacao_relatorio, ano_relatorio,
                str(linha_relatorio['Creditos']), '-', '-', '-', str(linha_relatorio['Turma'])
            ]), file=out_txt)
        else:
            encontrados += 1
            # Para cada disciplina correspondente encontrada em todas as matrizes
            for index_unificado, linha_unificado in busca_unificado.iterrows():

                # Se os créditos forem iguais ignora
                if linha_unificado['Créditos'] == linha_relatorio['Creditos']:
                    continue
                # Se os créditos forem diferentes
                creditos_inconsistentes += 1
                print(";".join([
                    'Créditos não correspondem', str(linha_relatorio['Cod Curso']), linha_relatorio['Nome Curso'],
                    disciplina_relatorio, abreviacao_relatorio, ano_relatorio, str(linha_relatorio.Creditos),
                    linha_unificado['Matriz'], str(linha_unificado['Créditos']), str(linha_unificado['Classificação']),
                    str(linha_relatorio['Turma'])
                ]), file=out_txt)

        porcentagem = int((index + 1) / total * 100)
        print(f"{porcentagem}% : Curso {curso_relatorio:5}, disciplina {disciplina_relatorio}")

    # Dados na finalização
    print(f"\nDisciplinas verificadas: {encontrados}")
    print(f"Inconsistências de créditos: {creditos_inconsistentes}")
    print(f"Disciplinas não encontradas: {nao_encontrados}")
    print(f"CRs não encontrados: {cursos_sem_abreviacao}")


if __name__ == '__main__':
    main()
