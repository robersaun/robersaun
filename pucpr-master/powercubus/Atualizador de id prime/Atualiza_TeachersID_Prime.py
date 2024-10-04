'''

1 - armazenar xml
2 - armazenar teachers.json(oq vem do powercubus via API)
3 - converter o name do xml para sem acentos e tudo maiusculo
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


def get_teachers(project_id, access_token):

    my_headers = {'Authorization': f'Bearer {access_token}'}
    teachers_id = requests.get(f'https://api.powercubus.com.br/v1/teachers/list?timetable_id={project_id}',
                            headers=my_headers)

    return teachers_id.json()





def update_ID(prime_xml, teachers_json, project_id, access_token):
    my_headers = {

        'Authorization': f'Bearer {access_token}',
        'Content-type': 'application/json',
        'Accept': 'application/json'
    }


    nomes_nao_encontrados =[]
    name_encontrado = []
    for child_xml in prime_xml.find('teachers'):
        name_xml = child_xml.get('name').split(';')[0]
        name_xml_convertido = remover_acentos(name_xml).upper()

        for teachers in teachers_json:
            if child_xml.get('partner_id') != None:
                id = child_xml.get('partner_id')
            elif child_xml.get('customfield1') != None:
                id = child_xml.get('customfield1')
            else:
                print('Partner_id ou External_id não encontrado')
                quit()
            #comparando o name do xml prime com o name do teacher no powercubus
            #se for verdade verificar o external_id com o partner_id
            if name_xml_convertido.upper() == teachers['name'].upper():
                name_encontrado.append(name_xml_convertido.upper())

                dict_valores ={
                                    "abreviation": teachers['abreviation'],
                                    "name": teachers['name'],
                                    "maximum_daily_lessons":teachers['maximum_daily_lessons'],
                                    "working_days": teachers['working_days'],
                                    "idletimes": teachers['idletimes'],
                                    "external_id": str(id),
                                    "id": teachers['id']
                            }
                valores_json = json.dumps(dict_valores)


                response = requests.put(f'https://api.powercubus.com.br/v1/teachers/update?'
                                        f'timetable_id={project_id}',
                                        headers=my_headers,data=valores_json)


                print(f'{teachers["name"]: <100} Atualização da ID retornou código '
                      f'{response.status_code} ({response.reason})')
            else:
                nomes_nao_encontrados.append(name_xml_convertido.upper())


    #escrevendo o log de erros para mostrar qual professor nao esta no powercubus

    aux=[]
    for i in nomes_nao_encontrados:
        if i not in name_encontrado:
         if i not in aux:
            f.write(f"### {i} professor não encontrado no PowerCubus\n")
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



    project_id = input('ID do Projeto: ')
    access_token = input('Token de acesso: ')

    teachers_json = get_teachers(project_id, access_token)

    update_ID(ID_prime_xml, teachers_json, project_id, access_token)


if __name__ == "__main__":
    f = open("log_TeacherIDS.txt", "w+", encoding='utf-8')
    main()
    f.close()
    print('Atualização finalizada')
