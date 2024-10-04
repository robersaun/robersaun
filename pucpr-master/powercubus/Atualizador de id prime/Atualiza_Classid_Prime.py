
'''

1 - armazenar xml
2 - armazenar class.json(oq vem do powercubus via API)
3 - converter o name do xml para siglas(para conseguir comparar os name corretamente),deixar tudo em caps lock e tirar os acentos
4 - comparar o name do xml com o name do powercubus se for igual seguir para o proximo passo
 4.1 - se for igual atualizar o external_id do powercubus para o partner_id do xml
 
 Desenvolvido por Matheus Rosa
 Última atualização: 16/05/2022

'''
import json
from lxml import etree
import xml.etree.ElementTree as ElementTree
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import requests
from unicodedata import normalize




def remover_acentos(txt):
    return normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')

def get_file_name():
    Tk().withdraw()
    file_name = askopenfilename(filetypes=[('XML', '.xml')])
    return file_name


def get_class(project_id, access_token):

    my_headers = {'Authorization': f'Bearer {access_token}'}
    class_id = requests.get(f'https://api.powercubus.com.br/v1/classes/list?timetable_id={project_id}',
                            headers=my_headers)

    return class_id.json()


def converter_name_class(name_class,dict_sigla):
    name_class_convertido = name_class
    for i in dict_sigla.items():
        if name_class == i[0]:
            name_class_convertido = i[1]

    return name_class_convertido


def update_ID(prime_xml, class_json, project_id, access_token, file_json_sigla):
    my_headers = {

        'Authorization': f'Bearer {access_token}',
        'Content-type': 'application/json',
        'Accept': 'application/json'
    }

    dict_sigla = dict(file_json_sigla)
    nomes_nao_encontrados =[]
    for child_xml in prime_xml.find('classes'):
        name_xml = child_xml.get('name').split(';')[0]
        name_xml_convertido = converter_name_class(name_xml,dict_sigla)

        for classes in class_json:

            #comparando o name do xml prime com o name da class no powercubus
            #se for verdade verificar as id

            if remover_acentos(name_xml_convertido.upper()) == classes['name'].split(' ')[1].upper():

                    dict_valores = {

                                    'grid_id': int(classes['grid_id']),
                                    'unit_id': int(classes['unit_id']),
                                    'name': str(classes['name']),
                                    'grade': str(classes['grade']),
                                    'course': str(classes['course']),
                                    'external_id': str(child_xml.get('id')),
                                    'id': int(classes['id'])

                                    }
                    valores_json = json.dumps(dict_valores)


                    response = requests.put(f'https://api.powercubus.com.br/v1/classes/update?'
                                            f'timetable_id={project_id}',
                                            headers=my_headers,data=valores_json)

                    print(f'{classes["name"]: <100} Atualização da ID retornou código '
                          f'{response.status_code} ({response.reason})')
            else:
                nomes_nao_encontrados.append(f"{name_xml}-{name_xml_convertido.upper()}")


    #escrevendo o log de erros para mostrar qual turma nao esta no powercubus
    aux=[]
    for i in nomes_nao_encontrados:
        if i not in aux:
            f.write(f"### {remover_acentos(i)} turma não encontrada no PowerCubus\n")
            aux.append(i)
def corrige_caracteres_especiais(file_name):
    # Lê o arquivo
    with open(file_name, "r", encoding="windows-1252") as file:
        file_data = file.read()

    # Substitui o símbolo que causa erro
    file_data = file_data.replace(" & ", " &amp; ")

    # Sobrescreve o arquivo
    with open(file_name, "w") as file:
        file.write(file_data)


def main():
    print("Selecione o arquivo XML")


    file_name_xml = get_file_name()
    parser = etree.XMLParser(recover=True)
    tree_xml = ElementTree.parse(file_name_xml,parser=parser)
    ID_prime_xml = tree_xml.getroot()

    with open('siglas.json', 'r', encoding="utf-8") as j:
        file_json_sigla = json.loads(j.read())


    project_id = input('ID do Projeto: ')
    access_token = input('Token de acesso: ')

    class_json = get_class(project_id, access_token)

    update_ID(ID_prime_xml, class_json, project_id, access_token,file_json_sigla)


if __name__ == "__main__":
    f = open("log_ClassIDS.txt", "w+", encoding='utf-8')
    main()
    f.close()
    print('Atualização finalizada')

