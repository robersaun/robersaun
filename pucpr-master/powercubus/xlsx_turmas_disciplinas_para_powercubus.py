""""
Script para realizar a conversão de um arquivo excel 
contendo as turmas e as disciplinas e de um arquivo XML das turmas, disciplinas e salas 
para o template utilizado no power cubus
    
Desenvolvido por Matheus Rosa e Fernando Dias
Última atualização: 18/05/2022
"""
import json
import random
import re
import string
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from xml.etree import ElementTree
import pandas
from unidecode import unidecode


def cria_abreviacao_professor(nome, dict_abreviacao):
    if nome == 'SEM PROFESSOR':
        return f'S. Prof ' \
               f'{random.choice(string.ascii_uppercase + string.digits)}' \
               f'{random.choice(string.ascii_uppercase + string.digits)}'
    if nome == 'nan':
        return f'S. Prof ' \
               f'{random.choice(string.ascii_uppercase + string.digits)}' \
               f'{random.choice(string.ascii_uppercase + string.digits)}'
    if nome in dict_abreviacao:
        return dict_abreviacao[nome]
    else:
        try:
            abreviacao = nome.split(' ')[0][0:8] + ' ' + nome.split(' ')[-1][0]

            while abreviacao in dict_abreviacao.values():
                abreviacao = abreviacao[:-1] + random.choice(string.ascii_uppercase)

            dict_abreviacao[nome] = abreviacao
        except:
            print('erro no nome do professor: ',nome,'')
            abreviacao = ''
        return abreviacao


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
def armazenaDadosXml(file_name_xml):
    tree_xml = ElementTree.parse(file_name_xml)
    root = tree_xml.getroot()

    def convertePeriod(period):

        if period == "1":
            valor = "7:05-7:50"
        elif period == "2":
            valor = "7:50-8:35"
        elif period == "3":
            valor = "8:35-9:20"
        elif period == "4":
            valor = "9:40-10:25"
        elif period == "5":
            valor = "10:25-11:10"
        elif period == "6":
            valor = "11:10-11:55"
        elif period == "7":
            valor = "11:55-12:40"
        elif period == "8":
            valor = "12:40-13:25"
        elif period == "9":
            valor = "13:25-14:10"
        elif period == "10":
            valor = "14:10-14:55"
        elif period == "11":
            valor = "15:15-16:00"
        elif period == "12":
            valor = "16:00-16:45"
        elif period == "13":
            valor = "16:45-17:30"
        elif period == "14":
            valor = "17:30-18:15"
        elif period == "15":
            valor = "18:15-19:00"
        elif period == "16":
            valor = "19:00-19:45"
        elif period == "17":
            valor = "19:45-20:30"
        elif period == "18":
            valor = "20:45-21:30"
        elif period == "19":
            valor = "21:30-22:15"
        elif period == "20":
            valor = "22:15-23:00"
        else:
            valor = ""

        return valor

    def convertDays(day):
        if day == "1111111" or day == "111111":
            valor = "Cada dia"

        elif day == "1000000" or day == "100000":
            valor = "Segunda-feira"

        elif day == "0100000" or day == "010000":
            valor = "Terça-feira"

        elif day == "0010000" or day == "001000":
            valor = "Quarta-feira"

        elif day == "0001000" or day == "000100":
            valor = "Quinta-feira"

        elif day == "0000100" or day == "000010":
            valor = "Sexta-feira"

        elif day == "0000010" or day == "000001":
            valor = "Sábado"

        elif day == "0000001":
            valor = "Horario Flutuante"
        else:
            valor = ""

        return valor

    dic_subject = {}
    for subject in root.iter("subject"):
        id = subject.get("id")
        name = subject.get("name")

        dic_subject[id] = name

    dic_class = {}
    for classes in root.iter("class"):
        id = classes.get("id")
        name = classes.get("name")
        short = classes.get("short")

        dic_class[id] = name,short

    dict_retorna_group = {}
    for groups in root.iter("group"):
        id = groups.get('id')
        name = groups.get('name')

        dict_retorna_group[id] = name

    dict_classrooms = {}
    for classrooms in root.iter("classroom"):
        id = classrooms.get("id")
        name = classrooms.get("name")
        short = classrooms.get("short")

        dict_classrooms[id] = name,short

    dict_cards = {}
    for cards in root.iter("card"):
        # parametros
        id = str(cards.get('lessonid'))
        period = cards.get('period')
        day = cards.get('days')

        valor = f"({convertePeriod(period)}/{convertDays(day)})"
        try:
            valor_dict = dict_cards[id]

            if valor_dict.__contains__(valor):
                continue
            else:
                dict_cards[id] += valor
        except:

            dict_cards[id] = valor

    valores = []
    for lessons in root.iter("lesson"):
        id = lessons.get('id')
        subjectid = lessons.get('subjectid')
        groupid = lessons.get('groupids')
        classid = lessons.get('classids')
        classroomsid = lessons.get('classroomids')

        name_subject = dic_subject[subjectid]

        cr_curso = name_subject.split(";")[0]

        quantidade_classroom = classroomsid.count(',')
        quantidade_classid = classid.count(',')

        posicao_sala = 0
        posicao_turma = 0


        try:

            for i in range(0, quantidade_classid+1):
                posicao_sala = 0
                for x in range(0, quantidade_classroom+1):
                    try:
                        if classroomsid == "":
                            continue
                        sala = str(dict_classrooms[classroomsid.split(',')[posicao_sala]]).split(',')[0]
                        abrev_sala = str(dict_classrooms[classroomsid.split(',')[posicao_sala]]).split(',')[1]

                        turma = str(dic_class[classid.split(',')[posicao_turma]]).replace(',', '/')
                        periodo = turma.split("/")[0].split(";")[1]

                    except KeyError:
                        f.write(f'id {id} nao teve o card encontrado')
                        continue
                    try:
                        dia_hora = dict_cards[id]
                    except KeyError:
                        f.write(f"id {id} não teve card encontrado")
                        continue
                    try:
                        group = dict_retorna_group[groupid.split(',')[0]]
                    except:
                        f.write('erro no group da lesson' + id + '\n')
                        group = '---'
                        continue


                    if group.replace(' ', '').upper() == 'TURMACOMPLETA' or group.replace(' ', '').upper() == 'ENTIRECLASS':
                        group = 'Teórico'

                    quantidade_dia = dia_hora.count("/")

                    try:
                        for posicao_sala in range(0, quantidade_dia):
                            dia = dia_hora.split(')')[posicao_sala].split('/')[1]
                            turma = turma.replace('Manhã e Tarde', 'I')
                            valor = f"{cr_curso}|{periodo}|{turma}|{sala}|{abrev_sala}|{dia_hora}|{name_subject}|{dia}|{group}"
                            valores.append(valor)
                    except:
                        print('')
                        continue
                posicao_sala += 1
                posicao_turma += 1
        except:
            print(classroomsid)
    return valores

def gerarRelatorio(file_name_xlsx, file_name_xml, file_json_sigla):
    siglas = dict(file_json_sigla)

    def converteSigla(turma):
        if turma in siglas.keys():
            return siglas[turma]
        else:
            return turma


    #transformando o caminho em um read de excel


    df = pandas.read_excel(file_name_xlsx)
    lista_xml = armazenaDadosXml(file_name_xml)
    dic_disciplina = {}
    dic_turma = {}
    dic_grupo = {}
    dic_sala = {}




    #tratando os dados nulos(nan) do excel
    df['Docente 2022.2'].fillna('Sem Professor', inplace=True)
    df['Nome da disciplina'].fillna('Sem Disciplinas', inplace=True)
    df['Turma'].fillna('Sem Turma', inplace=True)
    df['Cód. do Curso'].fillna('Sem Turma', inplace=True)




    # Gera uma sigla única para cada disciplina e cada professor
    dict_siglas = {}
    dict_abreviacao = {}
    for dados in df.iterrows():

        nome_disciplina = dados[1]['Nome da disciplina'].upper()
        nome_professor = dados[1]['Docente 2022.2'].upper()
        #criando siglas disciplinas
        if nome_disciplina not in dict_siglas:
            dict_siglas[nome_disciplina] = cria_sigla_disciplina(nome_disciplina, dict_siglas)
        # criando siglas professores
        if nome_professor not in dict_abreviacao:
            dict_abreviacao[nome_professor] = cria_abreviacao_professor(nome_professor, dict_abreviacao)



    #organuzando os dados
    dict_planilha={}

    for dados in df.iterrows():

        #tratamento por cota do formato da planilha que as linhas continuam sem os dados
        cod_curso = str(dados[1]['Cód. do Curso'])

        if cod_curso == "Sem Turma":
            break
        else:

            matriz = dados[1]['Matriz']
            periodo = dados[1]['Período']
            periodo = str(periodo).split('.')[0]

            turma = dados[1]['Turma']
            turno = dados[1]['Turno']
            nome_disciplina = dados[1]['Nome da disciplina'].upper()
            nome_disciplina_corrigido = nome_disciplina.replace(" ", "").replace('-',"")
            nome_disciplina_comparacao = unidecode(nome_disciplina_corrigido)

            nome_professor = dados[1]['Docente 2022.2'].upper()
            carga_horaria = dados[1]['Carga Horária']
            cod_disciplina = dados[1]['Cód. da Disciplina']

            #tratamento de dados para o formatado do cód de origem da aula
            try:
                 modulacao_matriz = int(dados[1]['Modulação pela Matriz'])
            except ValueError:
                print('erro para encontrar a modulção: ',dados)
                continue
            tipo_atividade = dados[1]['Tipo de Atividade']

            # Conversão da modulação_curso para string
            modulacao_curso = str(dados[1]['Modulação']).split('.')[0]

            #tratamento da modulação da matriz, modulação 60 = 1 e modulação 30 = 2
            if (modulacao_matriz / 60) % 60 == 1 :
                modulacao_matriz = '1'

            else:
                 modulacao_matriz = '2'

            #tratamento do tipo de atividade utilizando apenas a inicial ou parte do nome da atividade
            if tipo_atividade == 'Prático':
                tipo_atividade = 'P'
            elif tipo_atividade == 'Tutoria':
                tipo_atividade = 'Tut'
            else:
                tipo_atividade = 'Teórico'


            #gerando o cód. de origem de origem da aula

            #tuma completa não possui volores para modulação matriz e modulação curso
            if tipo_atividade == 'Teórico':
                cod_origem_aula = f"{tipo_atividade}"
            else:
                cod_origem_aula = f"{tipo_atividade}{modulacao_matriz}-{modulacao_curso}"




            # tratando dados
            try:
                nome_turma = matriz.split(' ')[0]
            except:
                print('erro em trazer o nome da turma: ', matriz)
                continue
            try:
                ano_turma = matriz.split(' ')[1]
            except:
                matriz = matriz.replace('-','/')
                ano_turma = matriz.split('/')[1]


            # tratamento turno
            if turno == "Manhã e Tarde":
                turno = 'I'
            else:
                turno = turno[0]

            turma_comparacao = f"{nome_turma} - {periodo}{turma} - {turno}"
            turma_corrigida = f"{nome_turma} - {periodo}{turma} - {turno} - {ano_turma}"

            d = 0


            dic_sala_aula = {}
            for x in dic_turma:
                if dic_turma[d] == turma_comparacao and dic_disciplina[d] == nome_disciplina_comparacao and dic_grupo[d] == cod_origem_aula:
                   dic_sala_aula[d] = dic_sala[d]
                   d += 1

                else:
                    dic_sala_aula[d] = ""
                    d += 1
            key = F"{turma_comparacao}{nome_disciplina_comparacao}{cod_origem_aula}"
            dict_planilha[key] = {
                'Turma': turma_corrigida,
                'Etapa': periodo,
                'Curso': nome_turma,
                'Nome da disciplina': nome_disciplina,
                'Sigla da disciplina': dict_siglas[nome_disciplina],
                'Nome do Professor': nome_professor,
                'Nome abreviado do professor': dict_abreviacao[nome_professor],
                'Aulas': carga_horaria,
                'Código de origem da disciplina': cod_disciplina,
                'Código de origem da aula': cod_origem_aula,
            }
    dic_retorna_sala = {}
    dict_dados_final = {}
    key_final = 0
    for x in lista_xml:
        # corrigir nome turma
        barra = r"'\'"
        turma = x.split('|')[2].split('/')[0].replace(barra, '').replace(barra, '').replace(barra, '').replace(barra, '').replace("'", '').replace('(', '')
        nome_turma = converteSigla(turma.split(";")[0])
        periodo_turma = turma.split(";")[1]
        tipo_turma = turma.split(";")[2]
        turno = turma.split(";")[3]
        if turno == "Manhã":
            turno = "M"
        elif turno == "Noite":
            turno = "N"
        elif turno == "Manhã e Tarde":
            turno = 'I'
        nome_turma_corrigido = f"{nome_turma} - {periodo_turma}{tipo_turma} - {turno}"


        sala_xml = x.split("|")[3].replace('(', '').replace(";", "-").replace("'", '')


        # corrigir nome disciplina
        try:
            nome_disciplina = x.split('|')[6].split(';')[2]
            nome_disciplina_corrigido = nome_disciplina.replace(' ', '').upper().replace('-','')
            nome_disciplina_corrigido = unidecode(nome_disciplina_corrigido)
        except:
            break

        # corrigindo grupo turma
        grupo_turma = x.split('|')[8]

        key_dados_xml = f"{nome_turma_corrigido}{nome_disciplina_corrigido}{grupo_turma}"
        dic_retorna_sala[key] = sala_xml


        try:
            dict_dados_final[key_final] = {
                'Turma': dict_planilha[key_dados_xml]["Turma"],
                'Etapa': dict_planilha[key_dados_xml]["Etapa"] ,
                'Curso': dict_planilha[key_dados_xml]["Curso"],
                'Nome da disciplina': dict_planilha[key_dados_xml]["Nome da disciplina"],
                'Sigla da disciplina':dict_planilha[key_dados_xml]["Sigla da disciplina"],
                'Area': '-',
                'Nome do Professor': dict_planilha[key_dados_xml]["Nome do Professor"],
                'Nome abreviado do professor': dict_planilha[key_dados_xml]["Nome abreviado do professor"],
                'Email do Professor': "",
                'Local de Aula': sala_xml ,
                'Aulas':dict_planilha[key_dados_xml]["Aulas"],
                'Máximo de aulas diárias': '',
                'Agrupamento de aulas': '1',
                'Máximo de dias de aula na semana': '',
                'Permitir aulas em dias consecutivos': '1',
                'Código de origem da turma': '',
                'Código de origem da disciplina': dict_planilha[key_dados_xml]["Código de origem da disciplina"],
                'Código de origem do professor': '',
                'Código de origem do local': '',
                'Código de origem da área': '',
                'Nome da unidade da turma': '',
                'Grupo de períodos': 'Graduação Presencial',
                'Nome da unidade do local': '',
                'Sigla da unidade da turma': '',
                'Sigla da unidade do local': '',
                'Código de origem da aula': dict_planilha[key_dados_xml]["Código de origem da aula"],
                'Código do agrupamento': '',
                'Horário Fixo': '',

            }
        except:
            print(key_dados_xml , 'não foi encontrada')
            continue
        key_final += 1
    new_df = pandas.DataFrame(data=dict_dados_final)

    new_df = (new_df.T)

    new_df.to_excel(file_name_xlsx.replace('.xlsx', ' PowerCubus.xlsx'), sheet_name='Template', index=False)


def main():

   #selecionando arquivos
    print('Selecione o arquivo turma disciplina')
    Tk().withdraw()
    file_name_xlsx = askopenfilename(filetypes=[('xlsx', '.xlsx')])


    print('Selecione o arquivo xml')
    Tk().withdraw()
    file_name_xml = askopenfilename(filetypes=[('xml', ".xml")])


    with open('siglas.json', 'r', encoding="utf-8") as j:
        file_json_sigla = json.loads(j.read())


    gerarRelatorio(file_name_xlsx, file_name_xml, file_json_sigla)

if __name__ == '__main__':
    main()
