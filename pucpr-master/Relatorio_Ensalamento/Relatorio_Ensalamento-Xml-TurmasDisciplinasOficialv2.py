
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas
from unidecode import unidecode
import numpy
import xml.etree.ElementTree as ElementTree
import json
import warnings

def armazenaDadosAscXml(file_name_asc_xml,file_json_turmasfic):
    tree_xml = ElementTree.parse(file_name_asc_xml)
    root = tree_xml.getroot()
    erro_log_turma=""
    erro_log_sala =""
    #converte period
    #retorna starttime/endtime
    def convertePeriod(period):

        if period == "1":
            valor="7:05-7:50"
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
            valor=""

        return valor

    def convertDays(day):
        if day == "1111111" or  day =="111111":
            valor="Cada dia"

        elif day == "1000000" or  day =="100000":
            valor = "Segunda-feira"

        elif day == "0100000" or  day =="010000":
            valor = "Terça-feira"

        elif day == "0010000" or  day =="001000":
            valor = "Quarta-feira"

        elif day == "0001000" or  day =="000100":
            valor = "Quinta-feira"

        elif day == "0000100" or  day =="000010":
            valor = "Sexta-feira"

        elif day == "0000010" or  day =="000001":
            valor = "Sábado"

        elif day == "0000001":
            valor = "Horario Flutuante"
        else:
            valor=""

        return valor

    #armazenar a id como key e o name como value
    dict_subject={}
    for subject in root.iter("subject"):
        # parametros
        id = subject.get('id')
        name = subject.get('name')

        dict_subject[id] = name

    # armazenar a id como key e o name,short como value
    dict_classrooms = {}
    for classrooms in root.iter("classroom"):
        #parametros
        id = classrooms.get('id')
        name = classrooms.get('name')
        short = classrooms.get('short')

        dict_classrooms[id] = name,short

    # armazenar a id como key e o name,short como value
    dict_class = {}
    for classes in root.iter("class"):
        # parametros
        id = classes.get('id')
        name = classes.get('name')
        short = classes.get('short')

        turmas_fic = dict(file_json_turmasfic)

        if name.__contains__('Engenharia;') or name.__contains__('Multicom'):
            cr_curso = short.split(";")[5]
            periodo = short.split(";")[2]
            tipo = short.split(";")[3]
            turno = short.split(";")[4]

            chave_turma_fic = f"{cr_curso}|{periodo}"

            if chave_turma_fic in turmas_fic.keys():
                name =  turmas_fic[chave_turma_fic]

                new_name = f"{name};{periodo};{tipo};{turno}"
                name = new_name
            else:
                f.write(f'#800 erro ao converter a multicom name ={name} , short={short}''\n')
                continue

        dict_class[id] = name, short


    # armazenar a id como key e o period,days como value
    #para puxar a id da lesson e saber qual o horario e dia daquela aula
    dict_cards = {}
    for cards in root.iter("card"):
        #parametros
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

    dict_retorna_group= {}
    for groups in root.iter("group"):
        id = groups.get('id')
        name = groups.get('name')

        dict_retorna_group[id] = name

    valores=[]
    for lessons in root.iter("lesson"):
        id = lessons.get('id')
        subjectid = lessons.get('subjectid')
        groupid = lessons.get('groupids')
        classid = lessons.get('classids').split(',')[0]


        classroomsid = lessons.get('classroomids')

        name_subject = dict_subject[subjectid]

        cr_curso = name_subject.split(';')[0]

        try:

            turma = str(dict_class[classid]).replace(',', '/')
        except:
            erro_log_turma +=f'#201 Turma nao identificada na aula: {id}\n'
            continue


        try:
            periodo = turma.split('/')[0].split(';')[1]
        except:
            erro_log_turma +=f'#200 Erro no formato do nome da turma: aula= {id},  turma= {turma} sala= {classroomsid} \n'
            continue
        quantidade = classroomsid.count(',')
        posicao_sala=0

        if quantidade == 0:
            quantidade =1
        else:
            quantidade +=1

        for i in range(0,quantidade):
            try:

                #if classroomsid == "":
                 #   continue

                sala = str(dict_classrooms[classroomsid.split(',')[posicao_sala]]).split(',')[0]
                abrev_sala = str(dict_classrooms[classroomsid.split(',')[posicao_sala]]).split(',')[1]
            except KeyError:
                sala = "Sem sala"
                abrev_sala = "-"
                f.write(f'id {id} nao teve a classroom encontrada, foi definida como sem sala''\n')


            try:
                dia_hora = dict_cards[id]

            except KeyError:
                dia_hora = "-/Flutuante"
                f.write(f'id {id} nao teve o card encontrado, foi definida como Flutuante''\n')


            try:
                group = dict_retorna_group[groupid.split(',')[0]]
            except:
                f.write('erro no group da lesson'+id+'\n')
                group ='---'

            quantidade_dia = dia_hora.count('/')

            if group.replace(' ','' ).upper() == 'TURMACOMPLETA' or group.replace(' ','' ).upper() == "ENTIRECLASS":
                group = '  '

            try:
                for posicao in range(0,quantidade_dia+1):
                    dia = dia_hora.split(')')[posicao].split('/')[1]
                    turma = turma.replace('Manhã e Tarde','I')
                    valor =f"{cr_curso}|{periodo}|{turma}|{sala}|{abrev_sala}|{dia_hora}|{name_subject}|{dia}|{group}"
                    valores.append(valor)
            except:
                posicao_sala += 1
                continue
            posicao_sala += 1

    f.write(erro_log_turma)
    return valores


def armazenaDadosCapacidade(file_name_capacidade):

    #armazena os valores das colunas importantes
    todos_valores=[]

    df = pandas.read_excel(file_name_capacidade,sheet_name="BASE DE CAPACIDADES CWB",header=2)

    posicao = 0
    teste = df.columns
    for dados in df.iterrows():

        abrev_sala = df['CÓD. ASC'][posicao].replace(' ','')
        capacidade = df['CAPACIDADE'][posicao]

        posicao+=1

        todos_valores.append(f"{abrev_sala}|{capacidade}")

    return todos_valores

def armazenaDadosTurmaDisciplina(file_name_turma_disciplina,file_json_sigla):
    siglas = dict(file_json_sigla)

    def converteSigla(turma):
        if turma in siglas.keys():
            return siglas[turma]
        else:
            return turma

    #armazena os valores das colunas importantes
    todos_valores=[]

    df = pandas.read_excel(file_name_turma_disciplina)
    posicao = 0

    for i in df.iterrows():

            #valores key
            cr_curso = df['CR Curso'][posicao]
            nome_curso = df['Curso'][posicao]
            periodo = df['Período'][posicao]


            nome_turma = df['Turma'][posicao]
            nome_turma = converteSigla(nome_turma)

 #           if nome_turma.count('-') < 3:
  #              Turno = df['Turno'][posicao]
   #             turno_letra = Turno[0]
    #            parte1 = nome_turma.split('-')[0]
     #           try:
      #              parte2 = nome_turma.split('-')[1]
       #         except:
        #            print(nome_turma)
         #           continue
          #      parte3 = nome_turma.split('-')[2]
           #     nome_turma = f"{parte1}-{parte2}- {turno_letra} -{parte3}"

            #valores nao utilizados como key
            nome_disciplina = df['Disciplina'][posicao]
            id_disciplina = df['ID'][posicao]
            divisao = df['Divisão'][posicao]
            dia = df['Data da Semana'][posicao]
            hora_inicio = df['Horário de Iní.'][posicao]
            hora_final = df['Horário de Tér'][posicao]
            agrupamento = df['Agrupamento'][posicao]
            turno = df['Turno'][posicao]
            cortes = df['Tem Corte'][posicao]
            qtd_alunos = df['Qtde de Alunos Matri.'][posicao]
            qtd_vagas = df['Nr Vagas Cadastradas']
            posicao += 1

            if divisao == ' ' or divisao == '  '  or divisao == '   ' or divisao == 'Teórico':
                divisao = '  '

            if cortes == 'Não':
                continue


            else:
                valores = f"{cr_curso}|{nome_curso}|{periodo}|{nome_turma}|{nome_disciplina}|{id_disciplina}|{divisao}|{dia}|{hora_inicio}|{hora_final}|{agrupamento}|{turno}|{qtd_alunos}|{qtd_vagas}"
                todos_valores.append(valores)


    return todos_valores

def gerarRelatorio(file_name_turma_disciplina,file_name_asc_xml,file_json_sigla,file_name_capacidade,file_json_turmasfic):
    siglas = dict(file_json_sigla)
    log_sala = {}

    warnings.simplefilter(action='ignore', category=UserWarning)

    def converteSigla(turma):
        if turma in siglas.keys():
            return siglas[turma]
        else:
            return turma

    # armazenando dados
    array_turmas_disciplinas = armazenaDadosTurmaDisciplina(file_name_turma_disciplina,file_json_sigla)
    array_asc_xml = armazenaDadosAscXml(file_name_asc_xml,file_json_turmasfic)
    array_capacidade = armazenaDadosCapacidade(file_name_capacidade)

    dict_capacidade ={}

    for i in array_capacidade:
        abrev_sala = i.split('|')[0]
        capacidade = i.split('|')[1]

        dict_capacidade[abrev_sala] = str(capacidade)


    dict_valores = {}
    dict_retorna_cr_disciplina = {}
    retorna_horario_inicio = {}
    horario_repetido = []
    for i in array_turmas_disciplinas:
        # valores chave
        cr_curso_aluno = i.split('|')[0]
        periodo = i.split('|')[2]
        nome_turma_final = i.split('|')[3].replace('- 2022/1', '').replace('2022/2', '').replace('2023/1', '').replace('2023/2', '')
        nome_turma = i.split('|')[3].replace('- 2022/1', '').replace('-', '').replace('2022/2', '').replace('2023/1', '').replace('2023/2', '')
        # valores
        nome_disciplina = unidecode(i.split('|')[4]).upper()
        id_disciplina = i.split('|')[5]
        divisao = i.split('|')[6]
        dia = i.split('|')[7]
        hora_inicio = i.split('|')[8]
        hora_final = i.split('|')[9]
        cr_disciplina = i.split('|')[5]
        nome_curso = i.split('|')[1]
        agrupamento = i.split('|')[10]
        turno = i.split('|')[11]
        qtd_alunos = i.split('|')[12]
        qtd_vagas = i.split('|')[12]

        chave_dia = f"{periodo},{nome_turma},{cr_disciplina}".replace(' ', '').replace('2022/2','').upper()
        chave = f"{periodo},{nome_turma},{cr_disciplina},{hora_inicio},{dia},{divisao}".replace(' ', '').upper()

        dict_valores[chave] = {
            'cr_curso' : cr_curso_aluno,
            'nome_disciplina': nome_disciplina,
            'Curso': nome_curso,
            'id_disciplina': id_disciplina,
            'divisao': divisao,
            'Nome_Turma': nome_turma_final,
            'dia': dia,
            'hora_inicio': hora_inicio,
            'hora_final': hora_final,
            'cr_disciplina': cr_disciplina,
            'Agrupamento': agrupamento,
            'Turno': turno,
            'qtd_alunos': qtd_alunos,
            'qtd_vagas': qtd_vagas

        }

        chave_horario_repetido = f"{chave}{hora_inicio}"
        try:
            if chave_horario_repetido not in horario_repetido:
                retorna_horario_inicio[chave_dia] += f"|{hora_inicio},{dia},{divisao}"
                horario_repetido.append(chave_horario_repetido)
        except:
            retorna_horario_inicio[chave_dia] = f"{hora_inicio},{dia},{divisao}"

        # retorna id da disciplina
        # - tratamento do nome da disciplina
        nome_disciplina_corrigido = nome_disciplina.split('-')[-1].replace(' ', '')
        chave_disciplina = f"{nome_disciplina_corrigido}{cr_curso_aluno}"
        dict_retorna_cr_disciplina[chave_disciplina] = cr_disciplina

    # usado para utilizar as chaves da planilha e retornar a sala
    dict_retorna_salaxml = {}
    dict_retorna_abrev_salaxml ={}
    key_final = 0
    valores_repetidos = []
    dict_relatorio_final = {}

    # criar dict para retornar a sala
    for dados in array_asc_xml:

        barra = r"'\'"

        periodo = dados.split('|')[1]
        # - tratamento nome turma
        turma = dados.split('|')[2].split('/')[0].replace(barra, '').replace(barra, '').replace(barra, '').replace(barra, '').replace("'", '')
        turma = turma[1:]

        turma_nome = converteSigla(turma.split(';')[0])
        abrev_turma = dados.split('|')[2].split('/')[1].replace(barra, '').replace(barra, '').replace(barra,'').replace(barra, '').replace("'", '').replace('(', '')
        nome_turma = f"{turma_nome}{periodo}{abrev_turma.split(';')[3]}{turma.split(';')[-1][0]}"
        divisao = dados.split('|')[8]
        dia = dados.split('|')[7]
        abrev_sala = dados.split('|')[4]
        abrev_sala = abrev_sala.replace(')', '').replace("'", "").replace(' ', '')
        try:
            nome_disciplina = dados.split('|')[6].split(';')[2]
        except:
            f.write(dados.split('|')[6] + 'name esta no formato errado')
            continue
        # valores
        sala = dados.split('|')[3].replace(barra, '').replace(barra, '').replace("'", '').replace('(', '')

        try:
            cr = dados.split('|')[6].split(';')[3]

            cr_curso = dados.split('|')[6].split(';')[0]
            try:
                cr_disciplina = int(cr)
            except:
                nome_disciplina_corrigido = nome_disciplina.split('-')[-1].replace(' ', '').upper()
                nome_disciplina_corrigido = unidecode(nome_disciplina_corrigido)
                chave_disciplina = f"{nome_disciplina_corrigido}{cr_curso}"
                cr_disciplina = dict_retorna_cr_disciplina[chave_disciplina]
                cr_disciplina = str(cr_disciplina)
        except:
            continue
        chave_sala = f"{periodo},{nome_turma},{cr_disciplina},{dia},{divisao}".replace(' ', '').upper()

        # retornar sala xml
        try:
            if not dict_retorna_salaxml[chave_sala].__contains__(sala):
                dict_retorna_salaxml[chave_sala] += f",{sala}"
        except:
            dict_retorna_salaxml[chave_sala] = sala

        try:
            if not dict_retorna_abrev_salaxml[sala].__contains__(abrev_sala):
                dict_retorna_abrev_salaxml[sala] += f",{abrev_sala}"
        except:
            dict_retorna_abrev_salaxml[sala] = abrev_sala

    log_sem_salas = {}
    log_sala_abrev={}
    log_sem_turmas_disciplinas ={}
    for dados in array_asc_xml:
            erro_repetido = []
            # barra invertida para poder retirar ela do texto
            barra = r"'\'"
            # valores key
            cr_curso_aluno = dados.split('|')[0].replace(' ', '')
            periodo = dados.split('|')[1]
            abrev_turma = dados.split('|')[2].split('/')[1].replace(barra, '').replace(barra, '').replace(barra,'').replace(barra, '').replace("'", '').replace('(', '')
            dia = dados.split('|')[7]
            try:
                nome_disciplina = dados.split('|')[6].split(';')[2]
            except:
                f.write(dados.split('|')[6]+ 'esta no formato errado')
                continue

            # - tratamento nome turma
            turma = dados.split('|')[2].split('/')[0].replace(barra, '').replace(barra, '').replace(barra, '').replace(barra, '').replace("'", '')
            turma = turma[1:]

            turma_nome = converteSigla(turma.split(';')[0])
            # turma utilizada para melhor entendimento
            turma_final = f"{turma_nome} - {periodo} - {abrev_turma.split(';')[3]} - {abrev_turma.split(';')[4]}"
            # turma utilizada como key
            nome_turma = f"{turma_nome}{periodo}{abrev_turma.split(';')[3]}{turma.split(';')[-1][0]}"
            try:
                cr = dados.split('|')[6].split(';')[3]

                cr_curso = dados.split('|')[6].split(';')[0]
                try:
                    cr_disciplina = int(cr)
                except:
                    nome_disciplina_corrigido = nome_disciplina.split('-')[-1].replace(' ', '').upper()
                    nome_disciplina_corrigido = unidecode(nome_disciplina_corrigido)
                    chave_disciplina = f"{nome_disciplina_corrigido}{cr_curso_aluno}"
                    cr_disciplina = dict_retorna_cr_disciplina[chave_disciplina]
                    cr_disciplina = int(cr_disciplina)

            except:
                f.write(dados.split('|')[6]+ ' não foi encontrado o cr'+'\n')
                continue

            # valores
            sala = dados.split('|')[3].replace(barra, '').replace(barra, '')

            # chave
            chave = f"{periodo},{nome_turma},{cr_disciplina}".replace(' ','').upper().replace("'",'')

            # armazena a lista de horarios de inicio da turma e disciplina para alocar na chave
            try:
                lista_horario_de_inicio_dia = retorna_horario_inicio[chave]
            except:
                try:
                    chave = f"{periodo},{nome_turma[:-1]},{cr_disciplina}".replace(' ', '').upper().replace("'", '')
                    lista_horario_de_inicio_dia = retorna_horario_inicio[chave]
                except:
                    log_sem_turmas_disciplinas[chave] = ''
                    continue

            quantidade_horarios = lista_horario_de_inicio_dia.count('|')
            posicao_horario = 0

            for x in range(0, quantidade_horarios + 1):

                hora_inicio = lista_horario_de_inicio_dia.split('|')[posicao_horario].split(',')[0]
                dia = lista_horario_de_inicio_dia.split('|')[posicao_horario].split(',')[1]
                divisao = lista_horario_de_inicio_dia.split('|')[posicao_horario].split(',')[2]

                chave_dia_hora = f"{periodo},{nome_turma},{cr_disciplina},{hora_inicio},{dia},{divisao}".replace(' ', '').upper()

                try:
                    dia = dict_valores[chave_dia_hora]['dia']
                    hora_inicio = dict_valores[chave_dia_hora]['hora_inicio']
                    hora_final = dict_valores[chave_dia_hora]['hora_final']
                except:
                    try:
                        chave_dia_hora_e = f"{periodo},{nome_turma[:-1]},{cr_disciplina},{hora_inicio},{dia},{divisao}".replace(' ','').upper()
                        dia = dict_valores[chave_dia_hora_e]['dia']
                        hora_inicio = dict_valores[chave_dia_hora_e]['hora_inicio']
                        hora_final = dict_valores[chave_dia_hora_e]['hora_final']
                    except:
                        posicao_horario += 1
                        f.write('Não foi encontrado o dia e hora da turma '+ chave_dia_hora+'\n')
                        continue


                chave_sala = f"{periodo},{nome_turma},{cr_disciplina},{dia},{divisao}".replace(' ', '').upper()
                try:
                    classroom = dict_retorna_salaxml[chave_sala].replace("'", '').replace('(', '')
                except:
                    try:
                        chave_sala_turmafic = f"{periodo},{nome_turma[:-1]},{cr_disciplina},{dia},{divisao}".replace(' ', '').upper()
                        classroom = dict_retorna_salaxml[chave_sala_turmafic].replace("'", '').replace('(', '')
                    except:
                        log_sem_salas[chave_sala] = ''
                        continue

                quantidade = classroom.count(',')
                posicao_sala = 0

                if quantidade == 0:
                    quantidade = 1
                else:
                    quantidade += 1

                for i in range(0, quantidade):
                    sala = classroom.split(',')[posicao_sala]


                    #buscando a abreviação da sala
                    try:
                        abrev_sala = dict_retorna_abrev_salaxml[sala]

                    except:
                        log_sala[sala] = ''


                        continue

                    #buscando a capacidade da sala
                    try:
                        capacidade = dict_capacidade[abrev_sala]
                    except:
                        capacidade = "0"
                        log_sala_abrev[abrev_sala] =''

                    try:
                        valores = dict_valores[chave_dia_hora]
                    except:
                        f.write('#600 dia e hora nao encontrado '+chave_dia_hora+'\n')
                        continue
                    # gerando dict que vai ser inserido no excel
                    dict_relatorio_final[key_final] = {
                        'CR Curso': valores['cr_curso'],
                        'Curso': valores['Curso'],
                        'Periodo': periodo,
                        'Turma': valores['Agrupamento'],
                        'Turma Prime': f"{valores['Nome_Turma']} 2022/2",
                        'Nome disciplina': nome_disciplina,
                        'Turno': valores['Turno'],
                        'Divisão': valores['divisao'],
                        'Id disciplina': cr_disciplina,
                        'Sala': sala,
                        'Sala Abreviação': abrev_sala,
                        'Capacidade da Sala': capacidade.replace('ECV','0').replace('nan','0'),
                        'Qtde de Alunos Matri.': valores['qtd_alunos'],
                        'Nr Vagas Cadastradas': valores['qtd_vagas'],
                        'Dia': dia,
                        'Hora inicio': hora_inicio,
                        'Hora final': hora_final
                    }

                    key_final += 1
                    valores_repetidos.append(chave_dia_hora)
                    posicao_sala +=1
                posicao_horario += 1

    for chave in log_sala.keys():
        f.write('#300 Erro ao buscar a abreviação da sala:'+chave+'\n')
    for chave in log_sala_abrev.keys():
        f.write('#300 Erro ao buscar capacidade da sala:' + chave + '\n')

    for chave in log_sem_salas.keys():
        f.write('#300 Não foi encontrada a sala: '+chave+'\n')
    for chave in log_sem_turmas_disciplinas.keys():
        f.write('#444 Não foi encontrada no excel turmas disciplinas: '+chave+'\n')



    Relatorio = {}
    valores_repetidos = []
    # tirar repetidos
    for dados in dict_relatorio_final.items():
        if dados[1]['Divisão'] == ' ' or dados[1]['Divisão'] == '  ' or dados[1]['Divisão'] == '   ':
            dados[1]['Divisão'] = 'Teórico'
        if dados[1] not in valores_repetidos:
            Relatorio[dados[0]] = dados[1]
            valores_repetidos.append(dados[1])

    df = pandas.DataFrame(data=Relatorio)

    df = (df.T)
    df = df.drop_duplicates()
    df.to_excel('Relatorio_Ensalamento-XML-TurmasDisciplinas.xlsx',index=False)


def main():

   #selecionando arquivos
    print('Selecione o arquivo turma disciplina')
    Tk().withdraw()
    file_name_turma_disciplina = askopenfilename(filetypes=[('xlsx', '.xlsx')])

    print('Selecione o arquivo XML ASC')
    Tk().withdraw()
    file_name_asc_xml = askopenfilename(filetypes=[('xml', '.xml')])


    print('Selecione o arquivo de capacidade salas')
    Tk().withdraw()
    file_name_capacidade = askopenfilename(filetypes=[('xlsx', '.xlsx')])
    print('carregando...')

    with open('siglas.json', 'r', encoding="utf-8") as j:
        file_json_sigla = json.loads(j.read())

    with open('turmasfic.json', 'r', encoding="utf-8") as j:
        file_json_turmasfic = json.loads(j.read())

    gerarRelatorio(file_name_turma_disciplina,file_name_asc_xml,file_json_sigla,file_name_capacidade,file_json_turmasfic)


if __name__ == '__main__':
    f = open(f"Log_Turmas.txt", "w+", encoding='utf8')
    main()
    f.close()
