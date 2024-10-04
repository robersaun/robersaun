'''
Script recebe um XML do Prime e realiza a retirada de informações do cursos e gera um novo arquivo XML

Desenvolvido por Vinicius Tozo
Última atualização: 12/07/2021

'''

from tkinter import Tk
from tkinter.filedialog import askopenfilename
from xml.etree import ElementTree


def corrige_caracteres_especiais(file_name):
    # Lê o arquivo
    with open(file_name, "r", encoding="windows-1252") as file:
        file_data = file.read()

    # Substitui o símbolo que causa erro
    file_data = file_data.replace(" & ", " &amp; ")

    # Sobrescreve o arquivo
    with open(file_name, "w") as file:
        file.write(file_data)


def extrair_xml_curso(arquivo_xml, cr_curso):
    # Lê o XML
    tree_xml = ElementTree.parse(arquivo_xml)
    root = tree_xml.getroot()

    active_lessons = []
    active_subjects = []
    active_teachers = []
    active_classrooms = []
    active_classes = []
    active_groups = []

    # Remove disciplinas de outros cursos
    for subject in root.find("subjects").findall("subject"):
        curso_disciplina = subject.get("name").split(";")[0]

        if curso_disciplina == cr_curso:
            active_subjects.append(subject.get("id"))
        else:
            root.find("subjects").remove(subject)

    if len(active_subjects) == 0:
        print("Aviso: Nenhum dado encontrado no CR")

    # Remove aulas de outros cursos
    for lesson in root.find("lessons").findall("lesson"):
        if lesson.get("subjectid") in active_subjects:
            active_lessons.append(lesson.get("id"))
        else:
            root.find("lessons").remove(lesson)

    # Remove cards de outros cursos
    for card in root.find("cards").findall("card"):
        if card.get("lessonid") not in active_lessons:
            root.find("cards").remove(card)

    # Procura os dados que estão ligados a uma aula, e remove o resto
    for lesson in root.iter("lesson"):
        active_classes += lesson.get("classids").split(",")
        active_classrooms += lesson.get("classroomids").split(",")
        active_teachers += lesson.get("teacherids").split(",")
        active_groups += lesson.get("groupids").split(",")

    for classes in root.find("classes").findall("class"):
        if classes.get("id") not in active_classes:
            root.find("classes").remove(classes)

    for classroom in root.find("classrooms").findall("classroom"):
        if classroom.get("id") not in active_classrooms:
            root.find("classrooms").remove(classroom)

    for teacher in root.find("teachers").findall("teacher"):
        if teacher.get("id") not in active_teachers:
            root.find("teachers").remove(teacher)

    for group in root.find("groups").findall("group"):
        if group.get("id") not in active_groups:
            root.find("groups").remove(group)

    return tree_xml


def main():
    # Recebe o arquivo XML no formato Prime
    print("Selecione o arquivo XML")
    Tk().withdraw()
    arquivo_xml = askopenfilename(filetypes=[("XML", ".xml")])
    corrige_caracteres_especiais(arquivo_xml)

    curso = input("Digite o CR do curso: ")

    # Converte o formato do XML do formato Prime para o formato Ponto
    xml_convertido = extrair_xml_curso(arquivo_xml, curso)

    # Salva o arquivo formatado
    xml_convertido.write(f"arquivo_curso_{curso}.xml", encoding="windows-1252")

    print("Arquivo gerado com sucesso!")


if __name__ == "__main__":
    main()
