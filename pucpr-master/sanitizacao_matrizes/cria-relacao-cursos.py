''''
Script recebe o arquivo excel Base Ouro e gera uma relação dos aulnos com as suas respectivas turmas e curso

Desenvolvido por Vinicius Tozo
Última atualização: 13/08/2021

'''

import pandas

def cria_dicionario():
    dicionario = {
        '2275': 'CICH', '2244': 'CAUR', '2303': 'CAUR', '2255': 'CAUR', '2231': 'CCJR', '2225': 'CCPP', '2232': 'CCPP',
        '2226': 'CCRP', '2256': 'CPVS', '2271': 'CDES',
        '2272': 'CDES', '2357': 'CDMD', '2354': 'CDPD', '2342': 'CDDG', '2356': 'CDGR',
        '2236': 'CJOR',
        '2224': 'CJOR', '2215': 'CMUS', '2283': 'CMPM', '2238': 'CRPU', '2308': 'CTEA', '2078': 'CAGR', '2079': 'CAGR',
        '2074': 'CBIT', '2360': 'CCBI', '2072': 'CGAS', '1060': 'CGAS', '2212': 'CBED',
        '2311': 'CBED',
        '2069': 'CENF', '2369': 'CENF', '2167': 'CFAR', '2066': 'CFIS',
        '2166': 'CFIS', '2076': 'CMVT', '2070': 'CNUT', '2064': 'CODO', '2063': 'CPSI',
        '2071': 'CPSI',
        '2022': 'CDIR', '2023': 'CDIR', '2268': 'CLBI', '2221': 'CSOC', '2206': 'CEDF', '2312': 'CEDF', '2202': 'CLFL',
        '2302': 'CLFL', '2304': 'CFIL', '2349': 'CLFS', '2314': 'CHIS', '2214': 'CHIS',
        '2301': 'CLPI', '2203': 'CLPI', '2249': 'CMAT', '2201': 'CPED', '2210': 'CPED',
        '2121': 'CQUI',
        '2350': 'CQUI', '2220': 'CSSC', '2319': 'CTEO', '2219': 'CTEO', '115': 'LAD', '2501': 'LAD', '111': 'LCC',
        '2502': 'LDI', '2518': 'LDI', '119': 'LPO', '111148': 'LES', '2530': 'LME', '121': 'LPS', '124': 'LPS',
        '114': 'LTO', '2610': 'GAD', '2611': 'GAD', '2605': 'GDI', '2606': 'GDI', '2603': 'GFL', '2614': 'GPS',
        '2615': 'GPS', '2295': 'CIS', '2062': 'CMED', '2028': 'CADM', '2029': 'CADM',
        '2239': 'CADI',
        '2229': 'CCOT', '2230': 'CCOT', '2036': 'CECO', '2035': 'CECO',
        '2293': 'CIBP', '2129': 'CMKT', '2128': 'CMKT',
        '2227': 'CTUR', '2287': 'CCBS', '2245': 'CCCO', '2345': 'CCCO', '2246': 'CEAM',
        '2346': 'CEAM',
        '2222': 'CEBI', '2223': 'CEBI', '2243': 'CECV', '2343': 'CECV', '2248': 'CECP', '2348': 'CECP', '2147': 'CEAL',
        '2158': 'CECA', '2159': 'CECA', '2261': 'CEMT', '2241': 'CEPD', '2341': 'CEPD', '2263': 'CESF', '2351': 'CEET',
        '2140': 'CEEL', '2252': 'CEMC', '2352': 'CEMC', '2353': 'CQUE', '2253': 'CQUE', '2344': 'CJDI', '2145': 'CJDI',
        '2250': 'CBSI', '2402': 'TAD', '2420': 'TAG', '2424': 'TAG', '2404': 'TBB', '2429': 'TCC', '2430': 'TDI',
        '2431': 'TDI', '2422': 'TPO', '2410': 'TMV', '112152': 'TMV', '112153': 'TPS', '2423': 'TPS', '2289': 'CTSI'
    }
    dataframe = pandas.read_excel("base_ouro_completa.xlsx")
    for index, linha in dataframe.iterrows():
        chave = str(int(linha["CR Aluno"]))
        valor = str(linha["Turma Aluno"]).split(" - ")[0]
        dicionario[chave] = valor
    return dicionario


if __name__ == '__main__':
    print(cria_dicionario())
