import xml.etree.ElementTree as ElementTree
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas




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
    for classroom in root.iter("classroom"):
        new_name = classroom.get("name").replace("*", "").strip()
        classroom.set("name", new_name)
        new_short = classroom.get("short").replace("*", "").strip()
        classroom.set("short", new_short)


def fix_subjects(root,dicionario_curso,escola):
    cod_curso = []
    extras = ["AULA ESPECIALIZAÇÃO", "AULA OUTROS PRODUTOS"]
    for subject in root.iter("subject"):
        cod_curso.append(subject.get('name').split(';'))
        # Verifica a formatação
        if len(subject.get("name").split(";")) != 4:
            print(f"Aviso: Formato incorreto no nome da disciplina {subject.get('name')}")
            continue
        if len(subject.get("short").split(";")) != 1:
            print(f"Aviso: Formato incorreto na abreviação da disciplina {subject.get('short')}")
            continue

        # Adiciona o *1* para aulas da graduação
        #TODO-----------
        parts = subject.get("name").split(";")
        if parts[2][-1:] != "*" and parts[2] not in extras:
            parts[2] = parts[2] + "*1*"
            new_name = ";".join(parts)
            subject.set("name", new_name)

        subject_name = parts[2].replace("*1*", "").replace("*2*", "").replace("*3*", "")

        # Verifica o formato para as outras aulas
        if subject_name in extras and not (parts[2].__contains__("*2*") or parts[2].__contains__("*3*")):
            print(f"Aviso: disciplina '{parts[2]}' não contem o tipo de horário (*2* ou *3*) no nome")
        if subject_name in extras and not (parts[3].__contains__("AULESP2") or parts[3].__contains__("AULESP3")):
            print(f"Aviso: disciplina '{parts[2]}' não contem a indicação 'AULESP' no nome")
        if subject_name in extras and not subject.get("short").__contains__("*extra*"):
            print(f"Aviso: disciplina '{parts[2]}' não contem a indicação '*extra*' na abreviação")
    #Mostra os curso que são da escola escolhida mas não estão no XML, estão so na planilha
    x = 0
    for i in dicionario_curso:

        cod_dict = list(dicionario_curso)
        nome_dict = list(dicionario_curso.values())
        if cod_dict[x] not in cod_curso[0]:
            if escola == 'unificado':
                print('Aviso: o curso ', cod_dict[x], '-', nome_dict[x],' não esta no XML,portanto não tem disciplina')
            else:
                print('Aviso: o curso ', cod_dict[x],'-',nome_dict[x],'é da escola ',escola ,' e não esta no XML,portanto não tem disciplina')
        x += 1



def fix_classes(root):
    for classes in root.iter("class"):

        # Verifica a formatação
        if len(classes.get("name").split(";")) != 4:
            print(f"Aviso: Formato incorreto no nome da turma {classes.get('name')}")
        if len(classes.get("short").split(";")) != 6:
            print(f"Aviso: Formato incorreto na abreviação da turma {classes.get('short')}")

        new_name = classes.get("name").replace("Manhã e Tarde", "Integral").strip()
        classes.set("name", new_name)

        # Se o turno for integral também deve alterar o atributo "short"
        if new_name.split(";")[-1] == "Integral":
            new_name = classes.get("short").split(";")
            new_name[4] = new_name[4].replace("M", "I")
            new_name = ";".join(new_name)
            classes.set("short", new_name)


def fix_teachers(root, dicionario_escolas):
    for teacher in root.iter("teacher"):

        # Verifica a formatação
        if len(teacher.get("name").split(";")) != 2:
            print(f"Aviso: Formato incorreto no nome do professor {teacher.get('name')}")
        if len(teacher.get("short").split(";")) != 2:
            print(f"Aviso: Formato incorreto na abreviação do professor {teacher.get('short')}")

        matricula = teacher.get("name").split(";")[-1].strip()
        new_short = dicionario_escolas.get(matricula, "") + ";" + teacher.get("short").split(";")[-1]
        teacher.set("short", new_short)

        # Alerta se não encontra o professor no dicionário
        if matricula not in dicionario_escolas:
            print(f"Aviso: Professor não encontrado na base professor - escola: {matricula}")


def remove_dados_zerados(root):
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
    for subject in root.iter("subject"):
        all_subjects[subject.get("id")] = subject.get("name").split(';')
    all_idlessons =[]

    all_classes = {}
    for classes in root.iter("class"):
        all_classes[classes.get("id")] = classes.get("name")

    # Identifica o id do horário flutuante para excluir as aulas sem horário fixo
    id_horario_flutuante = ""
    for horario in root.iter("daysdef"):
        if horario.get("name") == "Horario Flutuante":
            id_horario_flutuante = horario.get("id")

    if id_horario_flutuante == "":
        print("Aviso: Horário Flutuante não encontrado, as aulas não serão removidas")

    # Cria lista de aulas ativasif all_subjects.get(lesson.get('subjectid')) not in extra:
    for card in root.iter("card"):
        active_lessons.append(card.get("lessonid"))
    #pegar id do subject
    id_subject = []
    id_subject_excluir = []
    id_lesson_excluir = []
    extra = ['AULA ESPECIALIZAÇÃO*1*', 'AULA ESPECIALIZAÇÃO*2*', 'AULA ESPECIALIZAÇÃO*3*', 'PERMANÊNCIA*1*',
             'PERMANÊNCIA EXTERNA*1*', 'AULA EXTERNA*1*']
    #pegando id que vao ser excluidos
    for lesson in root.iter("lesson"):
        if lesson.get("classroomids") == "":
            if all_subjects.get(lesson.get('subjectid'))[2] not in extra:
                id_subject_excluir.append(lesson.get('subjectid'))
                id_lesson_excluir.append(lesson.get('id'))
                print(f"Aviso: a disciplina '{all_subjects.get(lesson.get('subjectid'))},"
                      f"não tem sala alocada então vai ser editada para AULA REMOTA")
                lesson.set('groupids', '')
                lesson.set('classids', '')
        else:
          id_subject.append(lesson.get('subjectid'))

    # Apaga lessons que não estão em um card
    # TODO: Mensagem de confirmação com os dados da lesson antes de apagar (Turma, sala, disciplina e professor)
    for lesson in root.find("lessons").findall("lesson"):
        if lesson.get("id") not in active_lessons \
                or lesson.get("teacherids") == "" \
                or lesson.get("daysdefid") == id_horario_flutuante:
            root.find("lessons").remove(lesson)
    #pegar a id da lesson que esta no card
    card_idlesson_flutuante = []
    #pegar a id das aulas especialização
    id_aulaesp = []
    #pegar a id da lesson que é graduação e horario flutuante
    lesson_flu_graduação = []
    #pegar id do subject que é horario flutuante
    id_subject_flutuante = []
    #pegando a id da lesson que tem horario flutuante
    for card in root.find("cards").findall("card"):
        if card.get('days') == '0000001':
            card_idlesson_flutuante.append((card.get('lessonid')))
    id_subject_nexcluir = []
    #todo
    novo = ElementTree.Element('subject')
    novo.set('id', '0000000000000000')
    novo.set('name','9999;99;AULA REMOTA*1*;AULREM')
    novo.set('short', 'AULA REMOTA')
    novo.set('partner_id', '0000')
    root.find("subjects").insert(0,novo)



    for lesson in root.iter("lesson"):
        active_classes += lesson.get("classids").split(",")
        active_teachers += lesson.get("teacherids").split(",")
        active_groups += lesson.get("groupids").split(",")
        #armazenando o name das disciplinas que são especialização e permanecias
        aula_esp = ['9999;99;AULA ESPECIALIZAÇÃO*1*;AULESP1','9999;99;AULA ESPECIALIZAÇÃO*2*;AULESP2','9999;99;AULA ESPECIALIZAÇÃO*3*;AULESP3']
        # TODO Opção 2
        # pegar as disciplinas que não tem sala alocada e mudar para aula remota
        if lesson.get('id') in id_lesson_excluir:
            lesson.set('subjectid','0000000000000000')
        active_subjects.append(lesson.get("subjectid"))
        for subject in root.iter("subject"):
            name = subject.get('name')
            if name in aula_esp:
                id_aulaesp.append(subject.get('id'))
            #excluindo disciplinas do horario flutuante
            if subject.get('id') not in id_aulaesp:
                if subject.get('id') in id_subject_flutuante:
                    print(f"Aviso: a disciplina '{all_subjects.get(lesson.get('subjectid'))},"
                          f"é GRADUAÇÃO e tem HORARIO FLUTUANTE então vai ser excluida")
                    root.find("subjects").remove(subject)
         # pegando a id da lesson que tem horario flutuante e não é especialização
        if lesson.get('id') in card_idlesson_flutuante:
            if lesson.get('subjectid') not in id_aulaesp:
                lesson_flu_graduação.append(lesson.get('id'))
                id_subject_flutuante.append(lesson.get('subjectid'))

        # TODO excluir lessons que são horario flutuante
        for lessons in root.find("lessons").findall("lesson"):
            if lessons.get('id') in lesson_flu_graduação:
                root.find('lessons').remove(lessons)
            all_idlessons.append(lessons.get('id'))

    for lesson in root.iter("lesson"):
        active_classrooms += lesson.get("classroomids").split(",")

    #TODO remove card que é horario flutuante mas não é especialização
    # remover card que não tem lesson
    for cards in root.find("cards").findall("card"):
        if cards.get('lessonid') in lesson_flu_graduação or cards.get('lessonid') not in all_idlessons:
            root.find("cards").remove(cards)

    for classes in root.find("classes").findall("class"):
        if classes.get("id") not in active_classes:
            root.find("classes").remove(classes)

    for classroom in root.find("classrooms").findall("classroom"):
        if classroom.get('id') not in active_classrooms:
            root.find("classrooms").remove(classroom)

    for subject in root.find("subjects").findall("subject"):
        if subject.get("id") not in active_subjects:
            root.find("subjects").remove(subject)

    for teacher in root.find("teachers").findall("teacher"):
        if teacher.get("id") not in active_teachers:
            root.find("teachers").remove(teacher)

    for group in root.find("groups").findall("group"):
        if group.get("id") not in active_groups:
            root.find("groups").remove(group)
    #TODO mostrando os codigos de curso que existem no XML
    cod_curso = []
    for subject in root.iter("subject"):
        cod_curso.append(subject.get('name').split(';')[0])
    print('Aviso: Códigos de curso existente no XML: ', sorted(set(cod_curso)))


def converte_xml(file_name_xml, dicionario_escolas,dicionario_codCurso,escola):
    tree_xml = ElementTree.parse(file_name_xml)
    root = tree_xml.getroot()

    remove_dados_zerados(root)
    fix_daysdef(root)
    fix_classes(root)
    fix_teachers(root, dicionario_escolas)
    fix_subjects(root , dicionario_codCurso,escola)
    fix_classrooms(root)

    return tree_xml


def cria_dicionario_professor_escola():
    # Lê o arquivo excel que contêm os dados
    dataframe = pandas.read_excel("base professor - escola.xlsx", engine="openpyxl")

    # Cria um dicionário com os dados do excel
    dicionario = {}
    for index, row in dataframe.iterrows():
        dicionario[str(row["MATRÍCULA ADP"]).strip()] = str(row["ESCOLA DE VÍNCULO"]).strip()
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




def main():
    print('Informe o numero da opção desejada?')
    print('[1] Unificado')
    print('[2] Belas Artes')
    print('[3] Ciências da Vida')
    print('[4] Direito')
    print('[5] Educação a Distância')
    print('[6] Educação e Humanidades')
    print('[7] Medicina')
    print('[8] Negócios')
    print('[9] Politécnica')
    print('')
    resposta = input()


    if resposta == '1':
        resposta = 'unificado'
    elif resposta == '2':
        resposta = 'BELAS ARTES'
    elif resposta == '3':
        resposta ='Ciências da Vida'
    elif resposta == '4':
        resposta = 'Direito'
    elif resposta == '5':
        resposta = 'Educação a Distância'
    elif resposta == '6':
        resposta = 'Educação e Humanidades'
    elif resposta == '7':
        resposta = 'MEDICINA'
    elif resposta == '8':
        resposta = 'Negócios'
    elif resposta == '9':
        resposta = 'Politécnica'
    else:
        print('valor invalido, reinicie o conversor e escolha um valor valido')
        exit()

    # Cria dicionário com a relação da escola de vínculo de cada professor
    dicionario_de_escolas = cria_dicionario_professor_escola()
    dicionario_codCurso = cria_dicionario_escola(resposta)
    # Lê o arquivo XML no formato Prime
    print("Selecione o arquivo XML")
    Tk().withdraw()
    arquivo_xml = askopenfilename(filetypes=[("XML", ".xml")])
    corrige_caracteres_especiais(arquivo_xml)
    print(f"\tArquivo selecionado: {arquivo_xml}")

    # Converte o formato do XML do formato Prime para o formato Ponto
    xml_convertido = converte_xml(arquivo_xml, dicionario_de_escolas,dicionario_codCurso,resposta)

    # Salva o arquivo formatado
    xml_convertido.write("arquivo_gerado.xml", encoding="windows-1252")




if __name__ == "__main__":
    main()
    input("Arquivo gerado com sucesso!")
