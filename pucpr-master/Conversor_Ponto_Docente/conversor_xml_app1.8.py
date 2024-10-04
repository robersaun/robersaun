'''
Converte arquivo ASC XMl para que seja possível utilizá-lo para a realização do ponto, 
utiliza como base 2 planilhas que devem estar na mesma pasta do script
e o nome dessas planilhas devem estar iguais ao nome que está no código. 
ARQUIVO PONTO - 2021.2 v3.xlsx e base professor - escola.xlsx

Desenvolvido por Matheus Rosa
última Atualização: 09/05/2022
'''

import xml.etree.ElementTree as ElementTree
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import re
import PySimpleGUI
import pandas
import PySimpleGUI as psg

def corrige_caracteres_especiais(file_name):
    # Lê o arquivo
    with open(file_name, "r", encoding="windows-1252") as file:
        file_data = file.read()

    # Substitui o símbolo que causa erro
    file_data = file_data.replace(" & ", " &amp; ")

    # Sobrescreve o arquivo
    with open(file_name, "w") as file:
        file.write(file_data)


def fix_daysdef(root):
    for daysdef in root.iter("daysdef"):
        new_name = daysdef.get("name").replace("-feira", "").strip()
        daysdef.set("name", new_name)


def fix_classrooms(root):
    classroom_repetidas =[]
    for classroom in root.iter("classroom"):
        new_name = classroom.get("name").replace("*", "").strip()
        classroom.set("name", new_name)
        new_short = classroom.get("short").replace("*", "").strip()
        classroom.set("short", new_short)

        try:
            short = new_short.split('.')[2]
        except:
            f.write('Aviso[info]: a classroom '+'name='+new_name+',short='+new_short+' esta no formato errado\n')

        if classroom.get("name") not in classroom_repetidas:
            classroom_repetidas.append(classroom.get("name"))
        else:
            f.write(f"Aviso[info]: a classroom esta duplicada name= {classroom.get('name')}\n")

def fix_subjects(root,dicionario_curso,escola):
    cod_curso = []
    extras = ["AULA ESPECIALIZAÇÃO", "AULA OUTROS PRODUTOS"]
    for subject in root.iter("subject"):
        cod = subject.get('name').split(';')
        #pegando somente numeros
        cod_corrigido = re.findall(r'\d+',str(cod).split(',')[0])
        cod_corrigido = str(cod_corrigido).replace('[',' ').replace(']',' ').replace("'",' ').replace(' ','')
        if cod_corrigido not in cod_curso:
            cod_curso.append(str(cod_corrigido).replace('[',' ').replace(']',' ').replace("'",' ').replace(' ',''))
        # Verifica a formatação
        if len(subject.get("name").split(";")) != 4:
            f.write(f"Aviso[info]: Formato incorreto no nome da subject {subject.get('name')}\n")
            continue
        if len(subject.get("short").split(";")) != 1:
            f.write(f"Aviso[info]: Formato incorreto na abreviação da subject {subject.get('short')}\n")
            continue

        # Adiciona o *1* para aulas da graduação
        #TODO-----------
        parts = subject.get("name").split(";")
        if parts[2][-1:] != "*" and parts[2] not in extras:
            parts[2] = parts[2] + "*1*"
            new_name = ";".join(parts)
            subject.set("name", new_name)

        subject_name = parts[2].replace("*1*", "").replace("*2*", "").replace("*3*", "")
        #todo falta documentar
        # Verifica o formato para as outras aulas
        if subject_name in extras and not (parts[2].__contains__("*2*") or parts[2].__contains__("*3*")):
            f.write(f"Aviso[info]: a subject '{parts[2]}' não contem o tipo de horário (*2* ou *3*) no nome\n")
        if subject_name in extras and not (parts[3].__contains__("AULESP2") or parts[3].__contains__("AULESP3")):
            f.write(f"Aviso[info]: a subject '{parts[2]}' não contem a indicação 'AULESP' no nome\n")
        if subject_name in extras and not subject.get("short").__contains__("*extra*"):
            f.write(f"Aviso[info]: a subject '{parts[2]}' não contem a indicação '*extra*' na abreviação\n")
    #Mostra os curso que são da escola escolhida mas não estão no XML, estão so na planilha
    x = 0
    for i in dicionario_curso:

        cod_dict = list(dicionario_curso)
        nome_dict = list(dicionario_curso.values())

        if cod_dict[x] not in cod_curso:

            if escola == 'unificado':
                f.write('Aviso[info]: o curso '+ str(cod_dict[x])+ '-'+ str(nome_dict[x])+' não esta no XML,portanto não tem disciplina\n')
            else:
                f.write('Aviso[info]: o curso '+ str(cod_dict[x])+'-'+str(nome_dict[x])+' é da escola '+escola +' e não esta no XML,portanto não tem disciplina\n')

        x += 1



def fix_classes(root):
    class_repetidas =[]
    for classes in root.iter("class"):

        # Verifica a formatação
        if len(classes.get("name").split(";")) != 4:
            f.write(f"Aviso[info]: Formato incorreto no nome da class {classes.get('name')}\n")
        if len(classes.get("short").split(";")) != 6:
            f.write(f"Aviso[info]: Formato incorreto na abreviação da class {classes.get('short')}\n")

        new_name = classes.get("name").replace("Manhã e Tarde", "Integral").strip()
        classes.set("name", new_name)

        # Se o turno for integral também deve alterar o atributo "short"
        if new_name.split(";")[-1] == "Integral":
            new_name = classes.get("short").split(";")
            new_name[4] = new_name[4].replace("M", "I")
            new_name = ";".join(new_name)
            classes.set("short", new_name)
        if classes.get("name") not in class_repetidas:
            class_repetidas.append(classes.get("name"))
        else:
            f.write(f"Aviso[info]: a class esta duplicada name= {classes.get('name')}\n")


def fix_teachers(root, dicionario_escolas):
    for teacher in root.iter("teacher"):
        matricula = teacher.get("name").split(";")[-1].strip()
        # Verifica a formatação
        if len(teacher.get("name").split(";")) != 2:
            f.write(f"Aviso[info]: Formato incorreto no nome do professor {teacher.get('name')}\n")
        if len(teacher.get("short").split(";")) != 2:
            f.write(f"Aviso[info]: Formato incorreto na abreviação do professor {teacher.get('short')}\n")
            f.write(f"Aviso[ação]: Abreviação do professor: {teacher.get('name')} ,foi corrigida com a inclusão da escola\n")


        new_short = dicionario_escolas.get(matricula, "") + ";" + teacher.get("short").split(";")[-1]
        teacher.set("short", new_short)


        # Alerta se não encontra o professor no dicionário
        if matricula not in dicionario_escolas:
            f.write(f"Aviso[info]: Professor não encontrado na base professor - escola: {matricula}\n")


def remove_dados_zerados(root,escola):
    win['progbar'].update_bar(25)

    # Itens que não podem ser apagados
    active_lessons = []
    active_subjects = []
    active_teachers = []
    active_classrooms = []
    active_classes = []
    active_groups = []



    # Correspondencias de id - nome para mostrar nos alertas
    #todo
    all_subjects = {}
    #armazenar a id pra comparar com o class e excluir a class que nao tem a mesma id
    subject_crP= {}
    subject_crP_repetido = []
    for subject in root.iter("subject"):
        all_subjects[subject.get("id")] = subject.get("name").split(';')
        try:
            if subject.get('name').split(';')[0] not in subject_crP:
                subject_crP[subject.get('name').split(';')[0]] = subject.get('name').split(';')[1]
            else:
                string2 = '()[]'
                valor = subject_crP.get(subject.get('name').split(';')[0])
                var = valor,subject.get('name').split(';')[1]
                result = ','.join(char for char in var if char not in string2)
                subject_crP[subject.get('name').split(';')[0]] = result
        except:
            #todo falta documentar
            f.write(f'Aviso[ação]: a subject '+ subject.get('name')+' foi excluida pois o name esta no formato errado\n')
    all_idlessons =[]

    all_classes = {}
    for classes in root.iter("class"):
        all_classes[classes.get("id")] = classes.get("name")

    # Identifica o id do horário flutuante para excluir as aulas sem horário fixo
    id_horario_flutuante = ""
    for horario in root.iter("daysdef"):
        if horario.get("name") == "Horario Flutuante":
            id_horario_flutuante = horario.get("id")



    # Cria lista de aulas ativasif all_subjects.get(lesson.get('subjectid')) not in extra:
    for card in root.iter("card"):
        active_lessons.append(card.get("lessonid"))
    #pegar id do subject
    id_subject = []
    id_subject_excluir = []
    id_lesson_excluir = []
    extra = ['AULA ESPECIALIZAÇÃO*1*', 'AULA ESPECIALIZAÇÃO*2*', 'AULA ESPECIALIZAÇÃO*3*', 'PERMANÊNCIA*1*',
             'PERMANÊNCIA EXTERNA*1*', 'AULA EXTERNA*1*','PERMANÊNCIA ONLINE*1*']
    #pegando id que vao ser excluidos
    id_lesson = []
    id_aulaesp =[]

    #armazenando aulas com a cr 9999
    for subject in root.iter("subject"):

        if subject.get('name').split(';')[0] == '9999':
            id_aulaesp.append(subject.get('id'))

    for lesson in root.iter("lesson"):
        if lesson.get("classroomids") == "" and lesson.get('subjectid') not in id_aulaesp:
            id_lesson_excluir.append(lesson.get('id'))

    for lesson in root.iter("lesson"):
        #codigo que retirava as segundas posições quando havia mais de uma class
        #if ',' in lesson.get("classids"):
         #   new_classid = lesson.get("classids").split(',')[0]
          #  lesson.set('classids',new_classid)

        if lesson.get("classroomids") == ""and lesson.get('subjectid') not in id_aulaesp:
            if lesson.get("classids") == "":
                try:
                    if all_subjects.get(lesson.get('subjectid'))[0].replace("'","") != '9999':
                        if all_subjects.get(lesson.get('subjectid'))[2] != 'AULA EXTERNA*1*':
                            id_lesson.append(lesson.get('id'))
                            f.write(f"Aviso[info]: A subject {all_subjects.get(lesson.get('subjectid'))} esta sem Lesson e sem Class\n")
                except:
                    continue
                else:
                    id_subject_excluir.append(lesson.get('subjectid'))
                    id_lesson_excluir.append(lesson.get('id'))
                    escrever = all_subjects.get(lesson.get('subjectid'))
                    lesson.set('groupids', '')
                    lesson.set('classids', '')
        else:
            id_subject.append(lesson.get('subjectid'))
    win['progbar'].update_bar(35)
    # Apaga lessons que não estão em um card
    # TODO: Mensagem de confirmação com os dados da lesson antes de apagar (Turma, sala, disciplina e professor)
    for lesson in root.find("lessons").findall("lesson"):
        if lesson.get("id") not in active_lessons \
                or lesson.get("teacherids") == "" \
                or lesson.get("daysdefid") == id_horario_flutuante and lesson.get("id") not in id_lesson:
            f.write(f'Aviso[ação]: a lesson ID={lesson.get("id")} foi excluida pois não esta em nenhum CARD\n')
            root.find("lessons").remove(lesson)
    #pegar a id da lesson que esta no card
    card_idlesson_flutuante = []
    #pegar a id das aulas especialização

    #pegar a id da lesson que é graduação e horario flutuante
    lesson_flu_graduação = []
    #pegar id do subject que é horario flutuante
    id_subject_flutuante = []
    #pegando a id da lesson que tem horario flutuante
    for card in root.find("cards").findall("card"):
        if card.get('days') == '0000001':
            card_idlesson_flutuante.append((card.get('lessonid')))

    win['progbar'].update_bar(45)
    for lesson in root.iter("lesson"):
        #armazenando o name das disciplinas que são especialização e permanecias
        aula_esp = ['9999;99;AULA ESPECIALIZAÇÃO*1*;AULESP1','9999;99;AULA ESPECIALIZAÇÃO*2*;AULESP2','9999;99;AULA ESPECIALIZAÇÃO*3*;AULESP3']

        # pegar as disciplinas que não tem sala alocada e mudar para AULA EXTERNA
        #todo
        if lesson.get('id') in id_lesson_excluir:


            lesson.set('subjectid','0000000000000000')
            lesson.set('classids', '')
            lesson.set('groupids', '')
            f.write(f"Aviso[ação]: a lesson ID='{lesson.get('id')},"
                    f"não tem sala alocada então vai ser editada para AULA EXTERNA\n")

        for subject in root.iter("subject"):


            #excluindo disciplinas do horario flutuante
            if subject.get('id') not in id_aulaesp:
                if subject.get('id') in id_subject_flutuante:
                    f.write(f"Aviso[ação]: a subject '{all_subjects.get(lesson.get('subjectid'))},"
                            f"é GRADUAÇÃO e tem HORARIO FLUTUANTE então vai ser excluida\n")
                    root.find("subjects").remove(subject)
                #todo
                if subject.get('id') in id_subject_excluir:
                    f.write(f"Aviso[ação]: a subject '{subject.get('name')},"
                        f"não tem sala alocada então vai ser editada para AULA EXTERNA\n")

    for lesson in root.iter("lesson"):
         # pegando a id da lesson que tem horario flutuante e não é especialização
        if lesson.get('id') in card_idlesson_flutuante:
                if lesson.get('subjectid') not in id_aulaesp and lesson.get('subjectid') != '0000000000000000' :
                    lesson_flu_graduação.append(lesson.get('id'))

                    id_subject_flutuante.append(lesson.get('subjectid'))
                    f.write(f"Aviso[ação]: a subject '{all_subjects.get(lesson.get('subjectid'))},"
                            f"é GRADUAÇÃO e tem HORARIO FLUTUANTE então vai ser excluida\n")

    # TODO excluir lessons que são horario flutuante
    for lesson in root.find("lessons").findall("lesson"):
        if lesson.get('id') in lesson_flu_graduação:
            print(lesson.get('id'))
            root.find('lessons').remove(lesson)

        all_idlessons.append(lesson.get('id'))
    win['progbar'].update_bar(55)

    #criando a subject AULA EXTERNA
    novo = ElementTree.Element('subject')
    novo.set('id', '0000000000000000')
    novo.set('name', '9999;99;AULA EXTERNA*1*;AULEXT')
    novo.set('short', 'AULA EXTERNA')
    novo.set('partner_id', '0000')
    root.find("subjects").insert(0,novo)




    #Remove o asterisco no inicio do name da subject
    for subject in root.iter("subject"):
        try:
            if subject.get('name').split(';')[2] in extra:
                if subject.get('name').split(';')[0][:1] == "*":
                    new_subject = f"{subject.get('name').split(';')[0][1:]};{subject.get('name').split(';')[1]};" \
                                  f"{subject.get('name').split(';')[2]};{subject.get('name').split(';')[3]}"
                    subject.set('name', new_subject)
        except:
            continue

    win['progbar'].update_bar(65)
    #todo falta documentar
    for classes in root.find("classes").findall("class"):
        try:
            codigo = classes.get('short').split(';')[5]
            periodo = classes.get('short').split(';')[2]
            periodo_split = subject_crP.get(codigo)
            virgula = ','
            if periodo_split != None:
                if virgula in periodo_split:
                    periodo_split = subject_crP.get(codigo).split(',')
            if codigo not in subject_crP:
                 f.write('Aviso[info]: a Class '+'[ id = '+ classes.get('id')+'; short ='+classes.get('short')+']'+ ' não tem seu codigo associado a nenhuma subject\n')
            elif periodo not in periodo_split:
                 f.write('Aviso[info]: a Class '+'[ id = '+ classes.get('id')+'; short ='+classes.get('short')+']'+' não tem nenhuma subject associada a seu periodo\n')
        except:
                f.write(f"Aviso[info]: a Class {classes.get('short')} esta no formato errado\n")

    for lesson in root.iter("lesson"):
        active_subjects += lesson.get("subjectid").split(",")


    for lesson in root.iter("lesson"):

        active_teachers += lesson.get("teacherids").split(",")
        active_groups += lesson.get("groupids").split(",")
        active_classrooms += lesson.get("classroomids").split(",")

        count_classes = lesson.get("classids").count(',')

        if count_classes == 0:
            count_classes = 1
        else:
            count_classes+=1

        for posicao in range(0, count_classes):
            active_classes.append(lesson.get("classids").split(',')[posicao])


    # TODO remove card que é horario flutuante mas não é especialização
    # remover card que não tem lesson
    for cards in root.find("cards").findall("card"):
        if cards.get('lessonid') in lesson_flu_graduação or cards.get('lessonid') not in all_idlessons:
            f.write(
                f'Aviso[ação]: o card ID={cards.get("lessonid")} period={cards.get("period")} weeks={cards.get("weeks")} terms={cards.get("terms")} days={cards.get("days")}'
                f' foi excluido pois não tem nenhuma lesson\n')
            root.find("cards").remove(cards)

    for classes in root.find("classes").findall("class"):
        if classes.get("id") not in active_classes:
            f.write(f'Aviso[ação]: a class {classes.get("name")} foi excluida pois não tem nenhuma lesson\n')
            root.find("classes").remove(classes)

    for classroom in root.find("classrooms").findall("classroom"):
        if classroom.get('id') not in active_classrooms:
            f.write(f'Aviso[ação]: a classroom ID={classroom.get("id")} name={classroom.get("name")} foi excluida pois não tem nenhuma lesson\n')
            root.find("classrooms").remove(classroom)

    for subject in root.find("subjects").findall("subject"):
        if subject.get("id") not in active_subjects and subject.get('name').split(';')[0] != '9999':
            f.write(f'Aviso[ação]: a subject ID={subject.get("id")} name={subject.get("name")} foi excluida pois não tem nenhuma lesson\n')
            root.find("subjects").remove(subject)

    for subject in root.find("subjects").findall("subject"):
        if subject.get("id") not in active_subjects and subject.get('name').split(';')[2] == 'AULA EXTERNA*1*':
            root.find("subjects").remove(subject)

    for teacher in root.find("teachers").findall("teacher"):
        if teacher.get("id") not in active_teachers:
            f.write(f'Aviso[ação]: o teacher ID={teacher.get("id")} name={teacher.get("name")} foi excluido pois não tem nenhuma lesson\n')
            root.find("teachers").remove(teacher)

    for group in root.find("groups").findall("group"):
        if group.get("id") not in active_groups:
            f.write(f'Aviso[ação]: o group ID={group.get("id")} name={group.get("name")} foi excluido pois não tem nenhuma lesson\n')
            root.find("groups").remove(group)

    for group in root.find("groups").findall("group"):
        if group.get('classid') not in active_classes:
            f.write(
                f'Aviso[ação]: o group ID={group.get("id")} name={group.get("name")} foi excluido pois não tem nenhuma class\n')
            root.find("groups").remove(group)

    for cards in root.find("cards").findall("card"):
        if cards.get('lessonid') in id_lesson_excluir and cards.get('classroomids') != "":
            f.write(
                f'+ Aviso[info]: o card ID={cards.get("lessonid")} tem sala e tambem esta associado a aula externa \n')

    #mudando o short da sala para Ensalamento via Magister
    if escola == "BELAS ARTES":
        posicao = 0
    elif escola == "Ciências da Vida":
        posicao = 200
    elif escola == "Direito":
        posicao = 400
    elif escola == "Educação e Humanidades":
        posicao = 600
    elif escola == "MEDICINA":
        posicao = 800
    elif escola == "Negócios":
        posicao = 1000
    elif escola == "Politécnica":
        posicao = 1200
    else:
        posicao = 0
        exit()


    for classroom in root.find("classrooms").findall("classroom"):
        posicao+=1
        classroom.set('name', f'Ensalamento via Magister{posicao}')
        classroom.set('short', f'CWB.Ensalamento via.Magister{posicao}')


    #teste
    #novo = ElementTree.Element('classroom')
    #novo.set('id', '0000000000000000')
    #novo.set('name', 'Ensalamento via.Magister')
    #novo.set('short', 'CWB.Ensalamento via.Magister')
    #novo.set('partner_id', '0000')
    #root.find("classrooms").insert(0,novo)

    #for lesson in root.find("lessons").findall("lesson"):
    #    lesson.set('classroomids', '0000000000000000')

    #for card in root.find("cards").findall("card"):
    #    card.set('classroomids', '0000000000000000')


    win['progbar'].update_bar(85)

def converte_xml(file_name_xml, dicionario_escolas,dicionario_codCurso,escola):
    tree_xml = ElementTree.parse(file_name_xml)
    root = tree_xml.getroot()
    remove_dados_zerados(root,escola)
    fix_daysdef(root)
    fix_classes(root)
    fix_teachers(root, dicionario_escolas)
    fix_subjects(root, dicionario_codCurso,escola)
    fix_classrooms(root)
    return tree_xml


def cria_dicionario_professor_escola():
    # Lê o arquivo excel que contêm os dados
    dataframe = pandas.read_excel("base professor - escola.xlsx", engine="openpyxl")

    # Cria um dicionário com os dados do excel
    dicionario = {}
    for index, row in dataframe.iterrows():
        matricula = str(row["MATRÍCULA ADP"]).strip()
        #matricula = int(matricula)
        try:
            matricula = matricula.split('.')[0]
        except:
            continue

        dicionario[matricula] = str(row["ESCOLA DE VÍNCULO"]).strip()
    return dicionario

def cria_dicionario_escola(escola):
    # Lê o arquivo excel que contêm os dados
    dataframe = pandas.read_excel("ARQUIVO PONTO - 2021.2 v3.xlsx", engine="openpyxl",sheet_name="CURSOS")

    # Cria um dicionário com os dados do excel
    dicionario = {}
    dicionario_escola ={}
    if escola == 'unificado':
        for index, row in dataframe.iterrows():
            dicionario[str(row["Cod. Curso"]).strip()] = str(row["Nome Curso"]).strip()
        return dicionario
    else:
        for index, row in dataframe.iterrows():
            escola2 = str(row["Escola"]).strip()
            if escola == escola2:
                dicionario[str(row["Cod. Curso"]).strip()] = str(row["Nome Curso"]).strip()
        return dicionario


def main(arquivo_xml,escola):

    if escola == 'Unificado':
        resposta = 'unificado'

    elif escola == 'Belas Artes':
        resposta = 'BELAS ARTES'

    elif escola == 'Ciências da Vida':
        resposta ='Ciências da Vida'

    elif escola == 'Direito':
        resposta = 'Direito'

    elif escola == 'Educação a Distância':
        resposta = 'Educação a Distância'

    elif escola == 'Educação e Humanidades':
        resposta = 'Educação e Humanidades'

    elif escola == 'Medicina':
        resposta = 'MEDICINA'

    elif escola == 'Negócios':
        resposta = 'Negócios'

    elif escola == 'Politécnica':
        resposta = 'Politécnica'
    else:
        exit()

    # Cria dicionário com a relação da escola de vínculo de cada professor
    dicionario_de_escolas = cria_dicionario_professor_escola()
    dicionario_codCurso = cria_dicionario_escola(resposta)
    # Lê o arquivo XML no formato Prime
    win['progbar'].update_bar(10)


    nome_arquivo = str(arquivo_xml.split('/')[-1])
    print('carregando...')

    corrige_caracteres_especiais(arquivo_xml)

    print(f"\tArquivo selecionado: {arquivo_xml}",file=f)

    

    # Converte o formato do XML do formato Prime para o formato Ponto
    xml_convertido = converte_xml(arquivo_xml, dicionario_de_escolas,dicionario_codCurso,resposta)

    # Salva o arquivo formatado
    xml_convertido.write(f"{nome_arquivo}_Arquivo_gerado.xml", encoding="windows-1252")

    print("Log gerada com sucesso!")
    win['progbar'].update_bar(95)


if __name__ == "__main__":

    #define o tema da tela
    psg.theme('LightBrown13')

    layout = [
              [psg.Text(size=(70, 1))],
              [psg.Text('Selecione a escola', size=(30, 1), font='Arial 15', justification='center',
                        text_color='black')],
              [psg.Combo(['Belas Artes', 'Ciências da Vida', 'Direito', 'Educação a Distância',
                          'Educação e Humanidades', 'Medicina', 'Negócios', 'Politécnica'],
                         default_value='Belas Artes', key='Escola', readonly=True, size=(25, 1), pad=(43, 0))],
              [psg.Text('_' * 100)],
              [psg.Text('Selecione o XML', size=(30, 0), font='Arial 15', justification='center',
                        text_color='black', )],
              [psg.FileBrowse('Clique aqui para Selecionar', size=(0, 0), key='Arquivo',file_types=(("Arquivo XML", "*.xml"),), ), psg.Button('Confirmar',border_width=2)],
              [psg.Text('' ,visible=True, key='name', size=(100, 0), font='Arial 8',text_color='black')],
              [psg.Text('_' * 100)],
              [psg.ProgressBar(50, orientation='h', size=(30, 10), key='progbar', bar_color=['brown4', 'White'],
                               visible=False, border_width=1)],
              [psg.Text(size=(70, 1))],
              [psg.Button('Converter', disabled=True,border_width=2,size=(15,1))],
              [psg.Text('*Selecione um arquivo para habilitar o botão de conversão',size=(100, 0),key='aviso',font='Arial 8',text_color='black')]
              ]


    # Define Window
    win = psg.Window('Conversor XML Ponto', layout, size=(450, 390),element_justification='c')

    while True:
        # le os valores e os eventos
        event, values = win.read()

        if event is None:
            break

        elif event == 'Confirmar' and values['Arquivo'] != "":

            win['Converter'].update(disabled=False)
            win['aviso'].update(visible=False)
            win['name'].update(visible=True)
            win['name'].update('Arquivo Selecionado: ' + values['Arquivo'].split('/')[-1], font='Arial 8')

        elif event == 'Converter':

            win['Converter'].update(disabled=True)
            #definindo o parametro inicial da barra de carregamento
            win['progbar'].update_bar(0)
            #definindo a visibilidade da barra de carregamento
            win['progbar'].update(visible=True)

            #le os valores vindo da tela
            arquivo_xml = values['Arquivo']
            escola = values['Escola']

            #pega somente o nome do arquivo para que o log seja salvo na mesma pasta
            nome_arquivo = str(values['Arquivo'].split('/')[-1])
            f = open(f"{nome_arquivo}_Log.txt", "w+", encoding='utf8')

            main(values['Arquivo'],escola)

            f.close()
            win['progbar'].update_bar(100)
            win.close()
            print("Arquivo gerado com sucesso!")
