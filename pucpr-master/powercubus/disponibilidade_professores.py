"""Script para importação de disponibilidades de professores

Script para importação de disponibilidades de professores coletadas no Magister.
Recebe como entrada o arquivo XML exportado do Magister com as disponibilidades,
o id do projeto e o token de acesso ao sistema.
Acessa a API do PowerCubus para realizar as importações.

Utilizado pelo analista de TI no processo de elaboração de horários no PowerCubus.

Desenvolvido por Vinicius Tozo
Última atualização: 28/10/2021
"""
import json
import xml.etree.ElementTree as ElementTree
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import requests


def get_file_name():
    Tk().withdraw()
    file_name = askopenfilename(filetypes=[('XML', '.xml')])
    return file_name


def print_availability(root):
    for child in root.find('teachers'):
        nome = child.get('name').split(';')[0]
        abreviacao = child.get('short')

        disponibilidade = child.get('timeoff').replace('.', '').split(',')
        seg, ter, qua, qui, sex, sab = disponibilidade

        dias_da_semana = ['seg', 'ter', 'qua', 'qui', 'sex', 'sab']
        horas = '    07:05  07:50  08:35  09:40  10:25  11:10  11:55  12:40  13:25  14:10  15:15  16:00  16:45  ' \
                '17:30  18:15  19:00  19:45  20:45  21:30  22:15'

        print(nome)
        print(horas)

        for i in range(len(disponibilidade)):
            print(dias_da_semana[i], '  ' + '      '.join(disponibilidade[i]))
        print()


def get_teacher_availability(project_id, access_token, teacher_id):
    my_headers = {'Authorization': f'Bearer {access_token}'}
    availability = requests.get(f'https://api.powercubus.com.br/v1/teachers/list-availability?'
                                f'timetable_id={project_id}&teacher_id={teacher_id}', headers=my_headers)
    return availability.json()


def get_teachers(project_id, access_token):
    my_headers = {'Authorization': f'Bearer {access_token}'}
    teachers = requests.get(f'https://api.powercubus.com.br/v1/teachers/list?timetable_id={project_id}',
                            headers=my_headers)

    return teachers.json()


def update_availability(availability_xml, teachers_json, project_id, access_token):
    my_headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-type': 'application/json',
        'Accept': 'application/json'
    }

    # Posição de cada hora no arquivo de disponibilidades em xml
    hour_position = {'07:05': 0, '07:50': 1, '08:35': 2, '09:40': 3, '10:25': 4, '11:10': 5, '11:55': 6, '12:40': 7,
                     '13:25': 8, '14:10': 9, '15:15': 10, '16:00': 11, '16:45': 12, '17:30': 13, '18:15': 14,
                     '19:00': 15, '19:45': 16, '20:45': 17, '21:30': 18, '22:15': 19}

    # Para cada professor
    for child_xml in availability_xml.find('teachers'):
        # Nome sem o código do professor
        name_xml = child_xml.get('name').split(';')[0]
        # Disponibilidades, incluindo domingo com 100% de disponibilidade
        timeoff_xml = ['00000000000000000000'] + child_xml.get('timeoff') \
            .replace('.', '').replace('0', '2').replace('1', '0').split(',')

        # Busca o id do professor no powercubus
        powercubus_id = ''
        for teacher_json in teachers_json:
            if teacher_json['name'] == name_xml:
                powercubus_id = teacher_json['id']
                break

        # Se não estiver no powercubus passa para o próximo professor
        if powercubus_id == '':
            print(f'{name_xml: <100} Professor não está no PowerCubus')
            continue

        # Busca a disponibilidade no powercubus
        powercubus_availability = get_teacher_availability(project_id, access_token, powercubus_id)

        # Modifica a disponibilidade no json
        modificado = False
        for i in range(len(powercubus_availability)):
            new_value = int(
                timeoff_xml[powercubus_availability[i]['day_id'] - 1][
                    hour_position.get(powercubus_availability[i]['period']['start_time'])]
            )
            if new_value != powercubus_availability[i]['unavailable']:
                powercubus_availability[i]['unavailable'] = new_value
                modificado = True

        # Envia a disponibilidade atualizada para o powercubus se houver alterações
        if modificado:

            data = []
            for a in powercubus_availability:
                data.append({
                    "day_id": int(a['day_id']),
                    "period_id": int(a['period']['id']),
                    "unavailable": int(a['unavailable']),
                    "unavailable_resource": int(a['unavailable_resource']),
                })

            data = json.dumps(data)

            response = requests.put(f'https://api.powercubus.com.br/v1/teachers/update-availability?'
                                    f'timetable_id={project_id}&teacher_id={powercubus_id}',
                                    headers=my_headers, data=data)
            print(f'{name_xml: <100} Atualização da disponibilidade retornou código '
                  f'{response.status_code} ({response.reason})')
        else:
            print(f'{name_xml: <100} Disponibilidade já estava atualizada no PowerCubus')

    return


def main():
    print("Selecione o arquivo XML...\n")
    file_name_xml = get_file_name()
    tree_xml = ElementTree.parse(file_name_xml)
    availability_xml = tree_xml.getroot()

    print_availability(availability_xml)

    project_id = input('ID do Projeto: ')
    access_token = input('Token de acesso: ')

    teachers_json = get_teachers(project_id, access_token)

    update_availability(availability_xml, teachers_json, project_id, access_token)


if __name__ == "__main__":
    main()
