'''
Script cria um dicionário vinculando professores e as escolas e gera um arquivo XML para ser utilizado no ponto, para isso ele recebe um XML do Prime

Desenvolvido por Vinicius Tozo
Última atualização: 23/07/2021
'''


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


def fix_subjects(root):
    extras = ["AULA ESPECIALIZAÇÃO", "AULA OUTROS PRODUTOS"]
    for subject in root.iter("subject"):

        # Verifica a formatação
        if len(subject.get("name").split(";")) != 4:
            print(f"Aviso: Formato incorreto no nome da disciplina {subject.get('name')}")
        if len(subject.get("short").split(";")) != 1:
            print(f"Aviso: Formato incorreto na abreviação da disciplina {subject.get('short')}")

        # Adiciona o *1* para aulas da graduação
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
    all_subjects = {}
    for subject in root.iter("subject"):
        all_subjects[subject.get("id")] = subject.get("name")

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

    # Cria lista de aulas ativas
    for card in root.iter("card"):
        active_lessons.append(card.get("lessonid"))

    # Apaga lessons que não estão em um card
    # TODO: Mensagem de confirmação com os dados da lesson antes de apagar (Turma, sala, disciplina e professor)
    for lesson in root.find("lessons").findall("lesson"):
        if lesson.get("id") not in active_lessons \
                or lesson.get("teacherids") == "" \
                or lesson.get("daysdefid") == id_horario_flutuante:
            root.find("lessons").remove(lesson)

    for lesson in root.iter("lesson"):
        active_classes += lesson.get("classids").split(",")
        active_classrooms += lesson.get("classroomids").split(",")
        active_teachers += lesson.get("teacherids").split(",")
        active_groups += lesson.get("groupids").split(",")
        active_subjects.append(lesson.get("subjectid"))

        if lesson.get("classroomids") == "":
            print(f"Aviso: Aula da disciplina '{all_subjects.get(lesson.get('subjectid'))}'"
                  f" na turma '{all_classes.get(lesson.get('classids'))}' não tem sala alocada")

    for classes in root.find("classes").findall("class"):
        if classes.get("id") not in active_classes:
            root.find("classes").remove(classes)

    for classroom in root.find("classrooms").findall("classroom"):
        if classroom.get("id") not in active_classrooms:
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


def converte_xml(file_name_xml, dicionario_escolas):
    tree_xml = ElementTree.parse(file_name_xml)
    root = tree_xml.getroot()

    remove_dados_zerados(root)
    fix_daysdef(root)
    fix_classes(root)
    fix_teachers(root, dicionario_escolas)
    fix_subjects(root)
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


def main():
    # Cria dicionário com a relação da escola de vínculo de cada professor
    dicionario_de_escolas = cria_dicionario_professor_escola()

    # Lê o arquivo XML no formato Prime
    print("Selecione o arquivo XML")
    Tk().withdraw()
    arquivo_xml = askopenfilename(filetypes=[("XML", ".xml")])
    corrige_caracteres_especiais(arquivo_xml)
    print(f"\tArquivo selecionado: {arquivo_xml}")

    # Converte o formato do XML do formato Prime para o formato Ponto
    xml_convertido = converte_xml(arquivo_xml, dicionario_de_escolas)

    # Salva o arquivo formatado
    xml_convertido.write("arquivo_gerado.xml", encoding="windows-1252")

    input("Arquivo gerado com sucesso!")


if __name__ == "__main__":
    main()
