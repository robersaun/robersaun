'''
Script realiza a conversão do arquivo gerado pelo Power Cubus no formato Json para um arquivo XML ASC, 
realizando as correções de id, dados duplicados e caracteres especiais.

Desenvolvido por Matheus Rosa
Última atualização: 16/05/20222

'''

import json
import re
from xml.dom import minidom
import xml.etree.ElementTree as ElementTree
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from console_progressbar import ProgressBar


def recebe_nome_do_arquivo(mensagem):
    Tk().withdraw()
    file_name = askopenfilename(initialdir="../Path/For/JSON_file",
                                filetypes=[("Json File", "*.json")],
                                title=mensagem
                                )
    return file_name

def get_file_name():
    Tk().withdraw()
    file_name = askopenfilename(filetypes=[('XML', '.xml')])
    return file_name

def escrever_xml(file, root,file_json_sigla,pb):
    def converter_period(name):
        if name == "07:05":
            period = '1'
        elif name == "07:50":
            period = '2'
        elif name == "08:35":
            period = '3'
        elif name == "09:40":
            period = '4'
        elif name == "10:25":
            period = '5'
        elif name == "11:10":
            period = '6'
        elif name == "11:55":
            period = '7'
        elif name == "12:40":
            period = '8'
        elif name == "13:25":
            period = '9'
        elif name == "14:10":
            period = '10'
        elif name == "15:15":
            period = '11'
        elif name == "16:00":
            period = '12'
        elif name == "16:45":
            period = '13'
        elif name == "17:30":
            period = '14'
        elif name == "18:15":
            period = '15'
        elif name == "19:00":
            period = '16'
        elif name == "19:45":
            period = '17'
        elif name == "20:45":
            period = '18'
        elif name == "21:30":
            period = '19'
        elif name == "22:15":
            period = '20'
        else:
            period = ''
        return period

    # todo tag periods
    periods = ElementTree.SubElement(root, "periods", options="canadd,export:silent",
                                     columns="period,name,short,starttime,endtime")

    ElementTree.SubElement(periods, "period", name="07:05", short="07:05", period="1", starttime="7:05", endtime="7:50")
    ElementTree.SubElement(periods, "period", name="07:50", short="07:50", period="2", starttime="7:50", endtime="8:35")
    ElementTree.SubElement(periods, "period", name="08:35", short="08:35", period="3", starttime="8:35", endtime="9:20")
    ElementTree.SubElement(periods, "period", name="09:40", short="09:40", period="4", starttime="9:40",endtime="10:25")
    ElementTree.SubElement(periods, "period", name="10:25", short="10:25", period="5", starttime="10:25",endtime="11:10")
    ElementTree.SubElement(periods, "period", name="11:10", short="11:10", period="6", starttime="11:10",endtime="11:55")
    ElementTree.SubElement(periods, "period", name="11:55", short="11:55", period="7", starttime="11:55",endtime="12:40")
    ElementTree.SubElement(periods, "period", name="12:40", short="12:40", period="8", starttime="12:40",endtime="13:25")
    ElementTree.SubElement(periods, "period", name="13:25", short="13:25", period="9", starttime="13:25",endtime="14:10")
    ElementTree.SubElement(periods, "period", name="14:10", short="14:10", period="10", starttime="14:10",endtime="14:55")
    ElementTree.SubElement(periods, "period", name="15:15", short="15:15", period="11", starttime="15:15",endtime="16:00")
    ElementTree.SubElement(periods, "period", name="16:00", short="16:00", period="12", starttime="16:00",endtime="16:45")
    ElementTree.SubElement(periods, "period", name="16:45", short="16:45", period="13", starttime="16:45",endtime="17:30")
    ElementTree.SubElement(periods, "period", name="17:30", short="17:30", period="14", starttime="17:30",endtime="18:15")
    ElementTree.SubElement(periods, "period", name="18:15", short="18:15", period="15", starttime="18:15",endtime="19:00")
    ElementTree.SubElement(periods, "period", name="19:00", short="19:00", period="16", starttime="19:00",endtime="19:45")
    ElementTree.SubElement(periods, "period", name="19:45", short="19:45", period="17", starttime="19:45",endtime="20:30")
    ElementTree.SubElement(periods, "period", name="20:45", short="20:45", period="18", starttime="20:45",endtime="21:30")
    ElementTree.SubElement(periods, "period", name="21:30", short="21:30", period="19", starttime="21:30",endtime="22:15")
    ElementTree.SubElement(periods, "period", name="22:15", short="22:15", period="20", starttime="22:15",endtime="23:00")


    # todo tag daysdef
    daysdefs = ElementTree.SubElement(root, "daysdefs", columns="id,days,name,short")
    ElementTree.SubElement(daysdefs, "daysdef", id="6B3F12A34D97076A", name="Qualquer dia", short="X",days="1000000,0100000,0010000,0001000,0000100,0000010,0000001")
    ElementTree.SubElement(daysdefs, "daysdef", id="805BDEBDA010C1F2", name="Cada dia", short="E", days="1111111")
    ElementTree.SubElement(daysdefs, "daysdef", id="A8F237B2F9E6B675", name="Segunda-feira ", short="Seg",days="1000000")
    ElementTree.SubElement(daysdefs, "daysdef", id="CAB97D0425BEC0DA", name="Terça-feira", short="Ter", days="0100000")
    ElementTree.SubElement(daysdefs, "daysdef", id="60850D9314471122", name="Quarta-feira ", short="Qua", days="0010000")
    ElementTree.SubElement(daysdefs, "daysdef", id="DDF4362F91A753AB", name="Quinta-feira ", short="Qui", days="0001000")
    ElementTree.SubElement(daysdefs, "daysdef", id="E87DBCE4E7616290", name="Sexta-feira ", short="Sex", days="0000100")
    ElementTree.SubElement(daysdefs, "daysdef", id="23651B657C8E30B9", name="Sábado ", short="Sáb", days="0000010")
    ElementTree.SubElement(daysdefs,"daysdef",id="0E5CEC8F57D233C2",name="Horário Flutuante",short="Flutuante",days="0000001")

    # todo tag weeksdef
    weeksdefs = ElementTree.SubElement(root, "weekdefs", options="canadd,export:silent", columns="id,weeks,name,short")
    ElementTree.SubElement(weeksdefs, "weekdef", id="4D751412E50EC8F8", name="Todas as semanas", short="Todas",
                           weeks="1")

    # todo tag termsdefs
    termsdefs = ElementTree.SubElement(root, "termsdefs", options="canadd,export:silent", columns="id,terms,name,short")
    ElementTree.SubElement(termsdefs, "termsdefs", id="402E718B0AFD874E", name="Todo o ano", short="ANO", terms="1")

    # todo tag subjects
    # armazenando id para nao repetir subjects
    id_subjects = []
    # escrevendo xml
    subjects = ElementTree.SubElement(root, "subjects", options="canadd,export:silent",
                                      columns="id,name,short,partner_id")
    for sub in file['allocations']:
        # verificando se é uma id repetida
        if json.dumps(sub['subject']['id']) not in id_subjects:
            # convertendo o name para a forma correta [codigo cr;periodo;name;partner_id]
            name = f"{sub['class']['name'].split('-')[0].replace(' ', '')};{sub['class']['name'].split('-')[2][1].replace(' ', '')};{sub['subject']['name']};{json.dumps(sub['subject']['external_id'])}"
            ElementTree.SubElement(subjects, "subject", id=json.dumps(sub['subject']['id']), name=name,
                                   short=sub['subject']['abreviation']
                                   , partner_id=json.dumps(sub['subject']['external_id']))
            id_subjects.append(json.dumps(sub['subject']['id']))
    pb.print_progress_bar(15)
    # todo tag teachers
    # todo arrumar external_id
    id_teachers = []
    teachers = ElementTree.SubElement(root, "teachers", options="canadd,export:silent",
                                      columns="id,name,short,gender,color,email,mobile,partner_id,firstname,lastname")
    for tea in file['allocations']:
        # verificando se é uma id repetida
        if json.dumps(tea['teacher']['id']) not in id_teachers:
            # convertendo o lastname para o formato correto [lastname;external_id]
            lastname = f"{json.dumps(tea['teacher']['name']).split(' ', 1)[1]};{tea['teacher']['external_id']}"
            # convertendo o short para o formato correto [name;external_id]
            short = f";{tea['teacher']['name']}"
            name = f"{tea['teacher']['name']};{tea['teacher']['external_id']}"
            ElementTree.SubElement(teachers, "teacher", id=json.dumps(tea['teacher']['id']),
                                   firstname=json.dumps(tea['teacher']['name']).split(' ')[0], lastname=lastname
                                   , name=name, short=short, gender="", color="", email="", mobile='', partner_id="")
            id_teachers.append(json.dumps(tea['teacher']['id']))

    pb.print_progress_bar(25)
    # todo tag buildings
    buildings = ElementTree.SubElement(root, "buildings", options="canadd,export:silent", columns="id,name,partner_id")

    # todo tag classrooms/ location
    id_classrooms = []
    classrooms = ElementTree.SubElement(root, "classrooms", options="canadd,export:silent",
                                        columns="id,name,short,capacity,buildingid,partner_id")
    for cla in file['allocations']:
        # verificando se é uma id repetida
        if json.dumps(cla['location']['id']) not in id_classrooms:
            # convertendo a sigla para o formato correto
            sigla = cla['location']['name'][:2]

            if sigla.upper() == "BL":

                name = cla['location']['name']
                # armazenando o periodo ex:[Bloco 5]
                numero = name[6]
                # serarar por ; e armazenar a ultima parte
                parte = cla['location']['name'].split(';', 1)[1]
                name_corrigido = f"CWB.{sigla.upper()}{numero}.{parte}"

                ElementTree.SubElement(classrooms, "classroom", id=json.dumps(cla['location']['id']),
                                       name=cla['location']['name'], short=name_corrigido, capacity="*", buildingid="*",
                                       partner_id="")
                id_classrooms.append(json.dumps(cla['location']['id']))
            else:
                ElementTree.SubElement(classrooms, "classroom", id=json.dumps(cla['location']['id']),
                                       name=cla['location']['name'], short=cla['location']['name'], capacity="*",
                                       buildingid="*", partner_id="")
                id_classrooms.append(json.dumps(cla['location']['id']))
    pb.print_progress_bar(25)
    # todo tag grades
    grades = ElementTree.SubElement(root, "grades", options="canadd,export:silent", columns="grade,name,short")
    for i in range(1, 21):
        ElementTree.SubElement(grades, "grade", name=f"Etapa {i}", short=f"Ano {i}", grade=f"{i}")

    # todo tag classes
    classes = ElementTree.SubElement(root, "classes", options="canadd,export:silent",columns="id,name,short,classroomids,teacherid,grade,partner_id")
    id_classes = []
    # armazenar id e periodo
    classes_dict = {}
    siglas = dict(file_json_sigla)
    for cl in file['allocations']:
        # verificando se é uma id repetida
        if json.dumps(cl['class']['id']) not in id_classes:
            # convertendo o name para o formato correto
            name = cl['class']['name'].replace(' ', '')
            sigla = name.split('-')[1]
            for i in siglas.items():
                sigla_key = i[0]
                sigla_value = i[1]

                if sigla == sigla_value:
                    sigla_corrigida = sigla_key
                    break

                else:
                    sigla_corrigida = sigla

            name_corrigido = f"{sigla_corrigida};{name.split('-')[2][0]};{name.split('-')[2][1:]};{name.split('-')[3]}"

            # convertendo o short para o formato correto
            short_corrigido = f"CWB;{sigla_corrigida};{name.split('-')[2][0]};{name.split('-')[2][1:]};{name.split('-')[3]};{name.split('-')[0]}"

            ElementTree.SubElement(classes, "class", id=json.dumps(cl['class']['id']), name=name_corrigido,
                                   short=short_corrigido, teacherid="", classroomids="",
                                   grade="", partner_id=json.dumps(cl['class']['id']))
            id_classes.append(json.dumps(cl['class']['id']))
            classes_dict[cl['class']['id']] = cl['class']['grade']
    pb.print_progress_bar(35)
    # todo tag groups

    # todo rever pois a regra do external id vai mudar
    def convertedivisao(external_id):

     external_id = external_id.upper()
     if external_id.__contains__(f'TC') or external_id.__contains__(f'TURMA COMPLETA'):
        divisao = f'Turma completa'

        return divisao

     for i in range(1,25):
        for y in range(1, 25):
            teste = f'P{i}-{y}'
            if external_id.__contains__(f'P{i}-{y}'):
                divisao = f'P{i}-{y}'

                return divisao
            elif external_id.__contains__(f'TUT{i}-{y}'):
               divisao = f'tut{i}-{y}'

               return divisao


    def converteStudentcount(divisao,sigla_curso):
        #dividir o total de alunos pela divisao(o primeiro numero da divisao)
        #pegar a sigla da class, para ver qual e o total de alunos a ser utilizado na hora de dividir

        #Padrao = 60
        #Medicina Curitiba(CMED) = 90
        #Maringa(MMED) = 50
        #Medicina Londrina(LMED) = 30

        #tratando sigla curso
        sigla_curso = sigla_curso.upper().replace(' ','')
        studentcount = 0
        total = 0


        if sigla_curso == "CMED":
            total = 90

        elif sigla_curso == "MMED":
            total = 50

        elif sigla_curso == "LMED":
            total = 30

        else:
            total = 60

        if divisao == "Turma completa":
            studentcount = total
        else:
            if divisao != None:
                divisao_corrigida = re.findall(r'\d+',divisao)
                numero_divisao = int(divisao_corrigida[0])
                studentcount = total/numero_divisao




        return int(studentcount)




    groups = ElementTree.SubElement(root, "groups", options="canadd,export:silent",
                                    columns="id,classid,name,entireclass,divisiontag,studentcount,studentids")
    # armazenar name e class id pora nao repetir
    groups_classid = {}
    id_groups = []
    # criando uma id para os groups
    id = 170000
    # pegando todos os class ids e nome da turma
    idclass_idgroups = {}
    # todo rever pois a regra do external id vai mudar
    for gro in file['allocations']:
        id += 1
        #try:
        sigla_curso = gro['class']['name'].split('-')[1]

        nome = convertedivisao(gro['event']['external_id'])

        studentcount = converteStudentcount(nome,sigla_curso)
        classid = gro['event']['class_id']
        lessonid = gro['event']['id']
        externalid = gro['event']['external_id']


        id_groups.append(classid)
        groups_classid[id] = classid, nome,lessonid,studentcount,externalid

        # armazenando class id e id group pra usar na lesson



    # tirando items repetidos do groups_classid
    temp = []
    res = dict()
    for key, val in groups_classid.items():
        variavel = f"{val[0]},{val[1]},{val[2]}"
        if variavel not in temp:
            temp.append(variavel)
            res[key] = val
    # tirando items repetidos do idclass_idgroups

    # escrevendo os groups sem repetidos
    #todo rever pois a regra do external id vai mudar
    def converterDivisionTag(name):
        '''
        tabela division tag
        Turma completa         = 0
        Grupo1 / Grupo2/ P6    = 1
        Alunos / Alunas / Tut9 = 2
        P1 / Não Usar / Tut6   = 3
        P2 / Tut4 / P12 / P15  = 4
        '''
        name = str(name).upper()

        if name.__contains__('TC') or name.__contains__('TURMA COMPLETA'):
            divisiontag = "0"
        elif name[:2] == "P1" or name == "Não Usar" or name[:4] == "Tut6":
            divisiontag = "3"
        elif name == "Grupo 1" or name == "Grupo 2" or name[:2] == "P6":
            divisiontag = "1"
        elif name == "Alunos" or name == "Alunas" or name[:4] == "Tut9":
            divisiontag = "2"
        elif name[:2] == "P2" or name[:4] == "Tut4" or name[:4] == "P12" or name[:4] == "P15":
            divisiontag = "4"
        else:
            divisiontag = ""
        return divisiontag

    # groups devem ser imprimidos em ordem certa se nao interfere no arquivo final
    # exemplo tut9-1 deve ficar a frente do tut9-2

    for i in res.items():

        id = i[0]
        name = i[1][1]
        classid = str(i[1][0])
        studentcount = i[1][3]
        entireclass = 0

        if name.upper() == 'TURMA COMPLETA':
            entireclass = 0
        ElementTree.SubElement(groups, "group", id=str(id),
                                       name=name, classid=classid, studentids="",
                                       entireclass=str(entireclass), divisiontag=converterDivisionTag(name),
                                       studentcount=str(studentcount))

    pb.print_progress_bar(47)
    # todo tag students
    students = ElementTree.SubElement(root, "students", options="canadd,export:silent",
                                      columns="id,classid,name,number,email,mobile,partner_id,firstname,lastname")

    # todo tag studentsubjects
    studentsubjects = ElementTree.SubElement(root, "studentsubjects", options="canadd,export:silent",
                                             columns="studentid,subjectid,seminargroup,importance,alternatefor")

    # todo tag lesson


    def converterdaysdef(day_id):

        if day_id == '1':
            daysdef_corrigido = "0000001"
        elif day_id == '2':
            daysdef_corrigido = "100000"
        elif day_id == '3':
            daysdef_corrigido = "010000"
        elif day_id == '4':
            daysdef_corrigido = "001000"
        elif day_id == '5':
            daysdef_corrigido = "000100"
        elif day_id == '6':
            daysdef_corrigido = "000010"
        else:
            daysdef_corrigido = ""

        return daysdef_corrigido

    cards_idlessons_repetidos = []
    for car in file['allocations']:
        parametro = f"{json.dumps(car['event']['id'])},{converterdaysdef(json.dumps(car['day_id']))}"
        cards_idlessons_repetidos.append(parametro)
    # id das lessons
    id_lessons = []
    lessons = ElementTree.SubElement(root, "lessons", options="canadd,export:silent",
                                     columns="id,subjectid,classids,groupids,teacherids,classroomids,periodspercard,periodsperweek,daysdefid,weeksdefid,termsdefid,seminargroup,capacity,partner_id")
    x = 0

    lista_all_lessons = []
    id_lessons_corrigidas = []
    for les in file['allocations']:
        lista_all_lessons.append(f"{les['event']['id']},{les['event']['class_id']},{les['event']['subject_id']},{les['event']['lessons']},{les['event']['teacher_id']},{str(les['event']['location_id']).replace(' ','')},{str(les['event']['external_id']).replace(' ','')},{les['period']['start_time']},{les['day_id']},{json.dumps(les['event']['lessons'])}")



    lesson_lista = []
    lesson_lista_perweek = []
    pb.print_progress_bar(50)

    groups_lessons= {}
    groups_class ={}
    for i in groups_classid.items():
        nome = i[1][1]
        classid = i[1][0]
        lessonid = i[1][2]
        id = i[0]
        #nome;classid = id
        groups_class[f"{nome};{classid}"] = id
        #idlesson = id
        groups_lessons[int(lessonid)] = i[0]



    for les in lista_all_lessons:
            #variaveis lista
            classid = les.split(',')[1]
            subjectid = les.split(',')[2]
            teacherid = les.split(',')[4]
            classroomid = les.split(',')[5]
            period = converter_period(les.split(',')[7])
            dayid = converterdaysdef(les.split(',')[8])
            externalid = les.split(',')[6]
            lessonid = les.split(',')[0]
            aulas = les.split(',')[9]

            try:
                id_group = groups_lessons[int(les.split(',')[0])]
                #todo mudar groups
                #turma base
                base = les.split(',')[6]
                base = base.upper()
                if base.__contains__("BASE"):
                    turma_base = 1
                else:
                    turma_base = 0
                lesson_lista.append(
                    f"{classid};{turma_base},{subjectid},{teacherid},{classroomid},{id_group},{period},{dayid},{externalid},{lessonid},{aulas}")
                id_lessons.append(json.dumps(lessonid))
            except:

                continue



    # excluindo items repetidos da lista
    # removendo items duplicados de uma lista
    pb.print_progress_bar(56)
    lista_lesson_corrigida = []
    item_lesson_repetido = []
    for element in lesson_lista:
        item_part1 = element.split(',')[:5]
        item_part2 = element.split(',')[-2:][0]
        item = str(item_part1) + str(item_part2)
        if item not in item_lesson_repetido:
            item_lesson_repetido.append(item)
            lista_lesson_corrigida.append(element)

    # for i in lista_lesson_corrigida:
    # parametros utilizados: id_class,id_subject,,id_location,converterdaysdef,groups
    # lesson_lista_perweek.append(
    # f"{i.split(',')[0]},{i.split(',')[1]},{i.split(',')[3]},{i.split(',')[6]},{i.split(',')[7]}")

    # juntar groups que sao nomes iguais e ids diferentes
    juntar_groups_ids = []
    posicao_gro = 0
    juntar_groups_dict={}

    pb.print_progress_bar(58)



    def excluir_class_duplicados(id_class):
        barra = r"'/'"
        id_class = id_class.replace('/', ",").replace(barra, "")
        tamanho = id_class.count(',') + 1
        posicao = 0
        lista_repetidos = []
        id_corrigida = ""
        #retirando repetidos
        for i in range(0, tamanho):
            if id_class.split(',')[posicao] not in lista_repetidos:
                lista_repetidos.append(id_class.split(',')[posicao])
                id_corrigida += f"{id_class.split(',')[posicao]}/"
                id_corrigida = id_corrigida.replace(barra, '').replace("/", ",").replace("'", "")
            posicao += 1
        contagem = id_corrigida.count(';')
        posicao = 0
        new_id_corrigido = ""
        if str(contagem) == '1':
            return id_corrigida.split(';')[0]

        else:
        #ordena a turma base
            for i in range(0,contagem):
                turma_base = id_corrigida.split(',')[posicao].split(';')[1]
                id = id_corrigida.split(',')[posicao]
                if turma_base == "1":
                    new_id_corrigido = id.split(';')[0] +","+ new_id_corrigido
                else:
                    new_id_corrigido = new_id_corrigido +"," +id.split(';')[0]
                posicao+=1

            #parte1 = new_id_corrigido.split(',')[0]
            #parte2 = new_id_corrigido.split(',')[2]
            #new_id_corrigido = parte1+","+parte2
            if new_id_corrigido[0] == ",":
                new_id_corrigido = new_id_corrigido[1:]
            return new_id_corrigido

    # juntar classids na lesson
    pb.print_progress_bar(59)
    def excluir_professores_duplicados(id_teachers):
        id_teachers = id_teachers.replace('/',',')
        tamanho = id_teachers.count(',') + 1
        posicao = 0
        lista_repetidos = []
        id_corrigida = ""
        for i in range(0, tamanho):
            if id_teachers.split(',')[posicao] not in lista_repetidos:
                lista_repetidos.append(id_teachers.split(',')[posicao])
                id_corrigida += f"{id_teachers.split(',')[posicao]},"
            posicao += 1

        return id_corrigida[:-1]


    # juntando os groups ids
    dict_juntar_groups = {}
    for i in res.items():

        nome_group = i[1][1]
        id_group = i[0]
        classid = i[1][0]
        externalid = i[1][4].replace(' ','')


        chave = F"{nome_group}{classid}{externalid}"
        if nome_group != None:
            if chave in dict_juntar_groups:
                dict_juntar_groups[chave] += "/" + str(id_group)
            else:
                dict_juntar_groups[chave] = str(id_group)



    #juntando os classids
    dict_juntar_classid={}
    for x in lesson_lista:
        subjectid = x.split(',')[1]
        classroomid = x.split(',')[3]
        period = x.split(',')[5]
        teacherid = x.split(',')[2]
        dayid = x.split(',')[6]

        chave = f"{subjectid}{classroomid}{period}{teacherid}{dayid}"
        classid = x.split(',')[0]

        if chave in dict_juntar_classid:
            dict_juntar_classid[chave] += "/" + str(classid)
        else:
            dict_juntar_classid[chave] = str(classid)



    #juntando os teachers
    dict_juntar_teachers ={}
    for i in lesson_lista:
        #parametros
        subjectid = i.split(',')[1]
        classroomid = i.split(',')[3]
        starttime = i.split(',')[5]
        dayid = i.split(',')[6]
        externalid = i.split(',')[7]
        aulas = i.split(',')[9]

        #verificar parametros
        chave = f"{subjectid}{classroomid}{starttime}{dayid}{externalid}{aulas}"

        teacher_id = i.split(',')[2]

        if chave in dict_juntar_teachers:
            dict_juntar_teachers[chave] += "/" + str(teacher_id)
        else:
            dict_juntar_teachers[chave] = str(teacher_id)

    id_repetido = []
    pb.print_progress_bar(60)
    # escrevendo lessons que so tem um professor
    id_lessons_repetidos = []

    # escrever lessons
    for y in lesson_lista:

        # parametros
        classid = y.split(',')[0]
        subjectid = y.split(',')[1]
        classroomid = y.split(',')[3]
        starttime = y.split(',')[5]
        dayid = y.split(',')[6]
        externalid = y.split(',')[7]
        nome_group = convertedivisao(externalid)
        aulas = y.split(',')[9]
        teacherid = y.split(',')[2]
        period = y.split(',')[5]

        if nome_group != None:
            nome_group = convertedivisao(externalid)

        #chaves para acessar os dicts
        chave_class = f"{subjectid}{classroomid}{period}{teacherid}{dayid}"
        chave_groups = f"{nome_group}{classid.split(';')[0]}{externalid}"
        chave_teachers = f"{subjectid}{classroomid}{starttime}{dayid}{externalid}{aulas}"


        if y.split(',')[8] not in id_repetido:
            if y.split(',')[8] not in id_lessons_repetidos:

                ElementTree.SubElement(lessons, "lesson", id=str(y.split(',')[8]),
                                       classids=excluir_class_duplicados(dict_juntar_classid[chave_class]),
                                       subjectid=y.split(',')[1],
                                       periodspercard=y.split(',')[9],
                                       periodsperweek=f"{y.split(',')[9]}.0",
                                       teacherids=f"{excluir_professores_duplicados(dict_juntar_teachers[chave_teachers])}",
                                       classroomids=y.split(',')[3],
                                       groupids=dict_juntar_groups[chave_groups],
                                       capacity="*", seminargroup="",
                                       termsdefid="402E718B0AFD874E",
                                       weeksdefid="4D751412E50EC8F8",
                                       daysdefid="6B3F12A34D97076A", partner_id="")
                id_lessons_repetidos.append(y.split(',')[8])




    # todo tag cards
    # dict para nao repetir id e periods iguais
    cards_list = []
    cards = ElementTree.SubElement(root, "cards", options="canadd,export:silent",
                                   columns="lessonid,period,days,weeks,terms,classroomids")
    for car in file['allocations']:

        #day_id 1 == Domingo
        #day_id 2 == Segunda
        #day_id 3 == Terça
        #day_id 4 == Quarta
        #day_id 5 == Quinta
        #day_id 6 == Sexta
        #day_id 7 == Sábado

        #deixando os daydefs no formato correto
        if json.dumps(car['day_id']) == '1':
            daysdef_corrigido = "0000001"
        elif json.dumps(car['day_id']) == '2':
            daysdef_corrigido = "1000000"
        elif json.dumps(car['day_id']) == '3':
            daysdef_corrigido = "0100000"
        elif json.dumps(car['day_id']) == '4':
            daysdef_corrigido = "0010000"
        elif json.dumps(car['day_id']) == '5':
            daysdef_corrigido = "0001000"
        elif json.dumps(car['day_id']) == '6':
            daysdef_corrigido = "0000100"
        elif json.dumps(car['day_id']) == '7':
            daysdef_corrigido = "0000010"
        else:
            daysdef_corrigido = ""
        # armazenando todos os cards
        cards_list.append(
            f"{json.dumps(car['event']['id'])},{converter_period(car['period']['start_time'])},{daysdef_corrigido},{car['location']['id']}")
        # verficiar se existe um card com period e daydefs repetidos
    aux = []
    posicao = 0
    lista = list(cards_list)

    for i in lista:
        #if i.split(',')[1:] not in aux:
            ElementTree.SubElement(cards, "card", lessonid=i.split(',')[0],
                                   classroomids=i.split(',')[3],
                                   period=i.split(',')[1], weeks="1",
                                   terms="1",
                                   days=i.split(',')[2])
            aux.append(i.split(',')[1:])

    pb.print_progress_bar(80)
    return root

def corrige_caracteres_especiais(file_name):


    # Substitui o símbolo que causa erro
    file_data = file_name.replace(" & ", " &amp; ")
    file_data = file_name.replace("&quot;", "").replace('& amp;','').replace(" &amp; ", "").replace(" ", "")


    return file_data


def corrige_caracteres_especiais_prime(file_name):
    # Lê o arquivo
    with open(file_name, "r", encoding="windows-1252") as file:
        file_data = file.read()

    # Substitui o símbolo que causa erro
    file_data = file_data.replace(" & ", " &amp; ")

    # Sobrescreve o arquivo
    with open(file_name, "w") as file:
        file.write(file_data)

    return file_name

def atualiza_id_prime(xml_prime,xml_convertido):
    tree_xml = ElementTree.parse(xml_prime)
    root = tree_xml.getroot()



    dict_prime_name_cr ={}
    for sub in root.iter("subject"):
        cr = sub.get("name").split(";")[0]
        name = sub.get("name").split(";")[2]
        name = name.upper()
        partner_id = sub.get("partner_id")
        chave = f"{cr};{name}"
        if cr == "" or cr == None:
            print('A disciplina: id =',sub.get("id"),'esta com o cr vazio no prime')
        elif name == "" or name == None:
            print('A disciplina: id =', sub.get("id"), 'esta com o name vazio no prime')
        else:
            dict_prime_name_cr[chave] = partner_id

    for sub in xml_convertido.iter("subject"):
        cr = sub.get("name").split(";")[0]
        name = sub.get("name").split(";")[2]
        name = name.upper()
        chave = f"{cr};{name}"
        try:

            if dict_prime_name_cr[chave]:
                partner_id = dict_prime_name_cr[chave]

                sub.set("partner_id", partner_id)
                print('Atualização feita com sucesso! subject: ', name, 'foi atualizada!')
        except:
            print('subject: ',chave,'não foi encontrada no xml prime')

def main():
    pb = ProgressBar(total=100, prefix='Here', suffix='Now', decimals=3, length=50, fill='X', zfill='-')
    print("Selecione o arquivo JSON")
    # file_name_period = recebe_nome_do_arquivo('selecione o arquivo de periods')
    file_name_json = recebe_nome_do_arquivo('selecione o arquivo json')
   # armazenando o arquivo json
    with open(file_name_json, 'r', encoding="utf-8") as j:
        file_json = json.loads(j.read())
    with open('siglas.json', 'r', encoding="utf-8") as j:
        file_json_sigla = json.loads(j.read())

    print("Selecione o arquivo XML")
    Tk().withdraw()
    arquivo_xml_prime = askopenfilename(filetypes=[("XML", ".xml")])


    print('carregando...')

    root = ElementTree.Element("timetable", ascttversion="2021.6.2", importtype="database",
                               options="export:idprefix:%CHRID,import:idprefix:%TEMPID,groupstype1,decimalseparatordot,lessonsincludeclasseswithoutstudents,handlestudentsafterlessons",
                               defaultexport="1", displayname="aSc Timetables 2012 XML", displaycountries="")
    pb.print_progress_bar(5)
    tree = escrever_xml(file_json, root,file_json_sigla,pb)

    # Salva o arquivo formatado
    xml_string = minidom.parseString(ElementTree.tostring(tree, encoding="windows-1252")).toprettyxml(indent="   ")
    # forçar a mudança do enconding
    # contar quantas vezes o ? se repete para que na segunda vez sobreescreva o enconding correto
    posicao = 0
    contagem = 0
    lista = list(xml_string)
    for i in xml_string:
        if i == "?":
            contagem += 1
        if contagem == 2:
            contagem = 3
            lista[posicao] = "encoding='windows-1252'?"

        posicao += 1
    xml_string = "".join(lista)
    with open('arquivo_gerado_powercubus_para_xml.xml', "w", encoding="windows-1252") as f:
        xml_string = corrige_caracteres_especiais(xml_string)
        f.write(xml_string)

    arquivo_xml_prime = corrige_caracteres_especiais_prime(arquivo_xml_prime)
    atualiza_id_prime(arquivo_xml_prime,tree)
    pb.print_progress_bar(100)

if __name__ == "__main__":
    main()
    print('conversão finalizada.')
