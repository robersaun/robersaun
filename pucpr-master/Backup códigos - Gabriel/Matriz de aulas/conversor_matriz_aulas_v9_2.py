# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
from numpy import nan
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

# Ajusta as informações da tabela de aulas por período
def ajusta_info(df_aulas):
    cont = 1 # Contador do período
    linha = 0 # Contador de linha
    # Insere o número do período na coluna criada
    # Tira os valores desnecessários das linhas
    nome_coluna = [] # Cria uma lista contendo o nome de cada coluna
    for c in df_aulas.columns:
        nome_coluna.append(c)
    for dado in df_aulas['Ordem']:
        if (dado == f'{cont+1}º PERÍODO'):
            cont+=1
        if (str(dado)!='nan'):
            df_aulas['Período'][linha] = str(int(cont))
            if (str(dado)=='Ordem' or str(dado)=='TOTAL' or str(dado)==f'{cont}º PERÍODO'):
                df_aulas['Período'][linha] = nan
                coluna = 0
                while coluna < 11:
                    df_aulas[nome_coluna[coluna]][linha] = nan
                    coluna+=1
        linha+=1
    return df_aulas

# Verifica quais colunas possuem dados a serem validados
# Apaga as colunas desnecessárias/não nomeadas
def ajusta_colunas(df):
    colunas = []
    drop_colunas = []
    contador = 0
    for c in df.head():
        # Feito desse modo para assegurar que a coluna será encontrada
        if (c != f'Unnamed: {contador-1}' and 
            c != f'Unnamed: {contador}' and 
            c != f'Unnamed: {contador+1}'):
            colunas.append(str(c))
        else:
            drop_colunas.append(str(c))
        contador+=1
    for dc in range(len(drop_colunas)):
        df.drop(drop_colunas[dc],inplace=True,axis=1)
    return df

# Insere os dados conforme o template
def insere_dados(df_para, df_aulas):
    for col_aulas in df_aulas.head():
        for col_para in df_para.head():
            dados_col_aula = []
            cont_dados_aulas = 0
            if (' '+col_aulas == col_para or
                col_aulas == col_para or
                col_aulas+' ' == col_para or
                (col_aulas == 'CRED' and col_para == 'Crédito')):
                for dados_aula in df_aulas[col_aulas]:
                    dados_col_aula.append(str(dados_aula))
                    cont_dados_aulas+=1
                for i in range(len(dados_col_aula)):
                    try:
                        # Se o dado inserido puder ser transformado em int, faz a conversão
                        df_para[col_para][i] = int(dados_col_aula[i].split('.')[0])
                    except:
                        # Caso contrário, é armazenado como String
                        df_para[col_para][i] = dados_col_aula[i]
    return df_para

# Transforma os valores númericos, necessários, em int
# >> Apenas na tabela de aulas > TODO: Fazer um para a tabela de totais
# >> Antes da alteração da forma como os dados são lidos, deveria ocorrer com todos os dados
def converte_dtype(df_para):
    for col_para in df_para.head():
        dado_para = []
        if (col_para == 'AT' or
            col_para == 'AP' or
            col_para == 'TOTAL' or
            col_para == 'Crédito' or
            col_para == 'HA' or
            col_para == 'HR' or
            col_para == 'HR EAD' or
            col_para == 'CH (HR) Extensionista' or
            col_para == 'Número de Vagas' or
            col_para == 'Vagas Totais' or
            col_para == 'MT' or
            col_para == 'MP' or
            col_para == 'CH Docente Teórica' or
            col_para == 'CH Docente Prática' or
            col_para == 'CH Docente Total'):
            for dp in df_para[col_para]:
                dado_para.append(dp)
            for t_dp in range(len(dado_para)):
                try:
                    df_para[col_para][t_dp] = int(dado_para[t_dp])
                except:
                    continue
    return df_para

# Atentar a forma como a função é escrita: Deve ser feita em inglês
# Insere as funções na tabela de aulas
def ajusta_funcoes(df_para):
    inicio = 14
    fim = 1
    for col_para in df_para.head():
        if col_para == 'Período Matriz':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = '=$K$9'
        elif col_para == 'Escola / Campus':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = '=$I$9'
        elif col_para == 'Curso':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = '=$D$9'
        elif col_para == 'TOTAL':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=SUM(J{i+inicio}:K{i+inicio})'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUBTOTAL(109,L{inicio}:L{i+(inicio-1)})'
        elif col_para == 'Crédito':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=L{i+inicio}'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUBTOTAL(109,M{inicio}:M{i+(inicio-1)})'
        elif col_para == 'HA':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = F'=IF(AND(x{i+inicio}="Estágio",Y{i+inicio}="Externo"),O{i+inicio},L{i+inicio}*20)'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUBTOTAL(109,N{inicio}:N{i+(inicio-1)})'
        elif col_para == 'HR':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=L{i+inicio}*15'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUBTOTAL(109,O{inicio}:O{i+(inicio-1)})'
        elif col_para == 'HR EAD':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=IF(Z{i+inicio}="EAD",M{i+inicio},0)'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUBTOTAL(109,P{inicio}:P{i+(inicio-1)})'
        elif col_para == 'CH (HR) Extensionista': # TODO: Falta a referência correta
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = '=#REF!'
                    # f'=SUM(H{i+inicio}:I{i+inicio})'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = '=#REF!'
                    # f'=SUBTOTAL(109,Q{inicio}:Q{i+(inicio-1)})'
        elif col_para == 'Número de Vagas':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = '=M9'
        elif col_para == 'Vagas Totais':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=IFERROR(VLOOKUP(F{i+inicio},Lista!$AA$2:$AC$13,2,0),"-")'
        elif col_para == 'CH Docente Teórica':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=ROUNDUP(IF(ISBLANK(AA{i+inicio}),(IFERROR((S{i+inicio}/T{i+inicio})*J{i+inicio},0)),0),0)'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUBTOTAL(109,AC{inicio}:AC{i+(inicio-1)})'
        elif col_para == 'CH Docente Prática':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=ROUNDUP(IF(ISBLANK(AA{i+inicio}),(IFERROR((S{i+inicio}/U{i+inicio})*K{i+inicio},0)),0),0)'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUBTOTAL(109,AD{inicio}:AD{i+(inicio-1)})'
        elif col_para == 'CH Docente Total':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=SUM(AC{i+inicio}:AD{i+inicio})'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUBTOTAL(109,AE{inicio}:AE{i+(inicio-1)})'
    return df_para

# TODO: Inserir as funções da tabela de totais >> Não está funcionando
def ajusta_funcoes_t2(df_para):
    inicio = 117
    for col_para in df_para.head():
        if col_para == 'CRÉDITO':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) == 0):
                    df_para[col_para][i] = f'=D{inicio}'
        """elif col_para == 'HORA AULA':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = '=$K$9'
        elif col_para == 'HORA RELÓGIO':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = '=$K$9'"""
    return df_para

def validacao_dados(load_df_para):
    inicio = 14
    qtd_linhas = 100

    # Período matriz
    val_periodo_matriz = DataValidation(type="list", 
        operator="equal",
        formula1='"2021/1,2021/2,2022/1,2022/2,2023/1,2023/2,2024/1,2024/2,2025/1,2025/2"', 
        allow_blank=True)

    # Período
    val_periodo = DataValidation(type="list", 
        operator="equal",
        formula1='"1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16"', 
        allow_blank=True)

    # AT e AP
    val_at_ap = DataValidation(type="list", 
        operator="equal",
        formula1='"0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30"', 
        allow_blank=True)

    val_at_ap2 = DataValidation(type="whole", 
        operator="between",
        formula1=-1, 
        formula2=31,
        allow_blank=True)

    # Grupo da Disciplina
    val_g_disc = DataValidation(type="list", 
        operator="equal",
        formula1='"Curso,Escola,Institucional"',    
        allow_blank=True)

    # Tipo da Disciplina
    val_t_disc = DataValidation(type="list", 
        operator="equal",
        formula1='"Normal,Estágio,TCC,OMC"', 
        allow_blank=True)
    
    # Modalidade
    val_modalidade = DataValidation(type="list",
        operator="equal", 
        formula1='"EAD,Presencial"', 
        allow_blank=True)
    
    # Curso Pagante #TODO: Colocar aquela lista de cursos q tem no template
    # Curso Pagante >> Não funciona e tira a validação do resto da planilha >> Talvez pegar de outra parte da planilha?
    # val_pagante = DataValidation(type="list", 
    #    formula1=('"Administração,Administração - Internacional,Agronomia,Análise De Sistemas,Análise E Desenvolvimento De Sistemas,Arquitetura E Urbanismo,Arquitetura E Urbanismo Com Ênfase Internacional,Bacharelado Em Artes Visuais,Bacharelado Em Biologia,Bacharelado Em Ciências Sociais,Bacharelado Em Dança,'
    #    'Bacharelado Em Educação Física,Bacharelado Em Filosofia,Bacharelado Em História,Bacharelado Em Jogos Digitais,Bacharelado Em Letras Inglês Internacional,Bacharelado Em Teologia,Bacharelado Interdisciplinar Em Ciências E Humanidades,Bacharelado Interdisciplinar Em Saúde,Big Data E Inteligência Analítica,Biotecnologia,Cibersegurança,Ciência Da Computação,'
    #    'Ciências Biológicas - Bacharelado,Ciências Contábeis,Ciências Contábeis - Internacional,Ciências Da Religião (Licenciatura),Ciências Econômicas,Ciências Econômicas - Internacional,Cinema E Audiovisual,Comunicação Digital,Comunicação Social - Hab. Publicidade E Propaganda,Comunicação Social - Hab.: Jornalismo,Comunicação Social - Hab.: Relações Públicas,'
    #    'Desenho Industrial,Desenho Industrial - Design De Moda,Desenho Industrial - Hab.: Design Digital,Desenho Industrial - Hab.: Programação Visual,Desenho Industrial - Hab.: Projeto Do Produto,Design,Design De Moda,Design De Produto,Design Digital,Design Gráfico,Digital Data Marketing,Direito,Educação Física,Enfermagem,Engenharia,Engenharia Ambiental,'
    #    'Engenharia Biomédica,Engenharia Civil,Engenharia De Alimentos,Engenharia De Alimentos (Ênfase Em Agroindústria),Engenharia De Computação,Engenharia De Controle E Automação,Engenharia De Materiais E Nanotecnologia,Engenharia De Produção,Engenharia De Produção Mecânica,Engenharia De Redes De Comunicações,Engenharia De Software,'
    #    'Engenharia Elétrica - Eixos: Telecomunicações Eletrônica Ou Sistema De Potência E Energia,Engenharia Elétrica (Ênfase Em Telecomunicações),Engenharia Eletrônica,Engenharia Florestal,Engenharia Mecânica,Engenharia Mecatrônica,Engenharia Mecatrônica (Controle E Automação),Engenharia Química,Estética E Cosmética,Farmácia,Física,Fisioterapia,Fonoaudiologia,'
    #    'Gestão Comercial,Gestão Da Tecnologia Da Informação,Gestão De Marketing Em Mídias Digitais,Gestão De Recursos Humanos,Gestão Financeira,Gestão Pública,Inteligência Artificial Aplicada,Intercambio Curitiba,Interdisciplinar Em Negócios,International Business Program Ibp,Isoladas Curitiba,Jornalismo,Letras - Português E Inglês,Letras-Português-Inglês,'
    #    'Licenciatura Em Artes Visuais,Licenciatura Em Ciências Biológicas,Licenciatura Em Ciências Sociais,Licenciatura Em Dança,Licenciatura Em Educação Física,Licenciatura Em Filosofia,Licenciatura Em Física,Licenciatura Em História,Licenciatura Em Letras - Habilitação: Inglês,Licenciatura Em Letras-Português,Licenciatura Em Letras-Português-Espanhol,Licenciatura Em Matemática,'
    #    'Licenciatura Em Música,Licenciatura Em Química,Licenciatura Em Sociologia,Licenciatura Em Teatro,Logística,Marketing,Marketing - Internacional,Matemática,Medicina,Medicina Veterinária,Negócios Digitais,Nutrição,Odontologia,Pedagogia,Processos Gerenciais,Produção Musical,Psicologia,Química,Relações Públicas,Serviço Social,Sistemas De Informação,Summer Institute,'
    #    'Sup.De Formação De Prof°. P/ Ed. Infantil E Séries Iniciais Do Ensino Fundamental,Superior De Tecnologia Em Cidades Inteligentes,Superior De Tecnologia Em Ciências E Inovação Dos Alimentos,Superior De Tecnologia Em Design De Interiores,Superior De Tecnologia Em Gastronomia,Superior De Tecnologia Em Gestão Comercial,Superior De Tecnologia Em Gestão Da Produção Industrial,'
    #    'Superior De Tecnologia Em Gestão Da Qualidade,uperior De Tecnologia Em Gestão De Recursos Humanos,Superior De Tecnologia Em Gestão De Segurança Privada,Superior De Tecnologia Em Gestão Financeira,Superior De Tecnologia Em Jogos Digitais,Superior De Tecnologia Em Logística,Superior De Tecnologia Em Produção Digital Multiplataformas,Superior De Tecnologia Em Segurança Da Informação,'
    #    'Teatro,Tecnologia Em Análise E Desenvolvimento De Sistemas,Tecnologia Em Big Data E Inteligência Analítica,Tecnologia Em Gestão Comercial,Tecnologia Em Gestão Da Tecnologia Da Informação,Tecnologia Em Gestão De Marketing Em Mídias Digitais,Tecnologia Em Gestão De Recursos Humanos,Tecnologia Em Gestão Financeira,Tecnologia Em Gestão Hospitalar,Tecnologia Em Logística,'
    #    'Tecnologia Em Processos Gerenciais,Tradutor E Intérprete De Língua Espanhola"'), 
    #    allow_blank=True)

    # Adiconando no worksheet as validações
    load_df_para.add_data_validation(val_periodo_matriz)
    load_df_para.add_data_validation(val_periodo)
    load_df_para.add_data_validation(val_at_ap)
    load_df_para.add_data_validation(val_at_ap2)
    load_df_para.add_data_validation(val_g_disc)
    load_df_para.add_data_validation(val_t_disc)
    load_df_para.add_data_validation(val_modalidade)
    #load_df_para.add_data_validation(val_pagante)

    val_periodo_matriz.add('K9')
    val_periodo.add(f'F{inicio}:F{qtd_linhas+(inicio-1)}')
    val_at_ap.add(f'J{inicio}:K{qtd_linhas+(inicio-1)}')
    val_at_ap2.add(f'J{inicio}:K{qtd_linhas+(inicio-1)}')
    val_g_disc.add(f'W{inicio}:W{qtd_linhas+(inicio-1)}')
    val_t_disc.add(f'X{inicio}:X{qtd_linhas+(inicio-1)}')
    val_modalidade.add(f'Z{inicio}:Z{qtd_linhas+(inicio-1)}')
    #val_pagante.add(f'AA{inicio}:AA{qtd_linhas+(inicio-1)}')

    return load_df_para

# Os dados das tabelas são inseridos o template definido
def insere_template(load_df_para, load_df_para_2, load_template):
    fd_pagina_origem = load_df_para["Sheet1"]
    fd_pagina_origem_totais = load_df_para_2["Sheet1"]
    fd_pagina_destino = load_template['CRIAÇÃO DE MATRIZ PRESENCIAL']

    # Insere os dados da tabela de aulas no modelo
    for i in range(1, fd_pagina_origem.max_row+1):
        for j in range(1, fd_pagina_destino.max_column+1):
            try:
                if(i < 101):
                    fd_pagina_destino.cell(row=i+12, column=j+2).value = fd_pagina_origem.cell(row=i+12, column=j+2).value
                else:
                    break
            except:
                continue

    # Insere os dados da tabela de totais no modelo
    for i in range(1, fd_pagina_origem.max_row+1):
        for j in range(1, fd_pagina_destino.max_column+1):
            try:
                fd_pagina_destino.cell(row=i+115, column=j+3).value = fd_pagina_origem_totais.cell(row=i+115, column=j+3).value
            except:
                continue

    fd_pagina_destino = validacao_dados(fd_pagina_destino)

    load_template.save('matriz_template.xlsx') # A matriz inserida no template

# Função principal
def main():
    # Seleção de arquivo dos aulas
    print("Selecione a planilha de aulas")
    Tk().withdraw()
    planilha_aulas = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de aulas')
    # Mostra qual foi o arquivo selecionado das aulas
    print("Arquivo selecionado: " + str(planilha_aulas).split('/')[-1])
    
    # Selecionando a planilha PARA
    print("Selecione a planilha PARA")
    Tk().withdraw()
    planilha_para = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha PARA')
    # Mostra qual foi o arquivo selecionado PARA
    print("Arquivo selecionado: " + str(planilha_para).split('/')[-1])

    ''' -------------------------------------------------------------------------------
        Leitura das páginas do arquivo
        ------------------------------------------------------------------------------- '''
    
    print("Lendo arquivo")
    
    # Planilha de aulas é armazenado na variável
    df_aulas = pd.read_excel(planilha_aulas, dtype=str, header=1)

    # planilha PARA e armazenada em uma variavel, selecionando apartir do header 12
    df_para = pd.read_excel(planilha_para, dtype=str, header=12, skipfooter=5, index_col=1, sheet_name=0) # Seleciona os dados das aulas
    df_para2 = pd.read_excel(planilha_para, header=115, index_col=1, sheet_name=0) # Identifica os dados dos totais inidicados

    df_para = ajusta_colunas(df_para) # Faz os ajustes necessários nas colunas
    df_para2 = ajusta_colunas(df_para2)

    df_aulas.insert(loc=0, column='Período', value=nan) # Cria uma coluna para colocar o número do período
    df_aulas = ajusta_info(df_aulas) # Ajusta as informações e linhas da tabela de aulas
    
    # Remove os nan das células vazias
    df_aulas.dropna(inplace=True, how='all') # remove a linha, quando todas as células estiverem vazias
    df_aulas.drop_duplicates(inplace=True) # remove duplicadas, no caso, os espaços em branco e os títulos
    df_aulas = df_aulas.fillna("")
    
    df_para.drop_duplicates(inplace=True) # remove duplicadas, nesse caso, os espaços em branco e os títulos
    df_para = df_para.fillna("")
    
    df_para = insere_dados(df_para, df_aulas) # Pega os dados do dataframe de alunos e insere no template
    df_para = ajusta_funcoes(df_para) # Insere as funções da tabela 1 no local correto
    df_para2 = ajusta_funcoes_t2(df_para2)
    df_para = converte_dtype(df_para)

    '''------------------------------------------------------------------------------------------------'''
    
    df_para.to_excel('tabela_aulas_matriz.xlsx', index=False, startrow=12, startcol=2) # Cria a planilha utilizando o template PARA
    df_para2.to_excel('tabela_totais_matriz.xlsx', index=False, startrow=115, startcol=3) # Isso era para mexer na tabela 2

    load_df_para = load_workbook('tabela_aulas_matriz.xlsx') # Cria o arquivo com o nome do curso
    load_df_para_2 = load_workbook('tabela_totais_matriz.xlsx') # Cria o arquivo com os totais
    load_template = load_workbook(planilha_para)
    
    insere_template(load_df_para, load_df_para_2, load_template)

if __name__ == '__main__':
    # Chamada da função main
    main()
