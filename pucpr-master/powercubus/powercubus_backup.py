"""Script para fazer backup dos dados no PowerCubus

Script para backup de todos os dados que podem ser exportados do sistema via API.
Recebe como entrada o id do projeto e o token de acesso ao sistema e gera um arquivo json para cada tipo de dado.

Pode ser utilizado pelo analista de TI para auxiliar no processo de elaboração de horários no PowerCubus.

Desenvolvido por Vinicius Tozo
Última atualização: 03/02/2022
"""
import json
import os
from datetime import datetime

import requests


def main():
    project_id = input('ID do Projeto: ')
    access_token = input('Token de acesso: ')

    backup(project_id, access_token)


def backup(project_id, access_token):
    my_headers = {'Authorization': f'Bearer {access_token}'}

    data_objects = ['classes', 'events', 'grids', 'groups', 'locations', 'periods',
                    'processes', 'subjects', 'teachers', 'timetables', 'units']

    folder = f'PowerCubus Backup {datetime.now().strftime("%Y-%m-%d %H%M")}'

    if not os.path.exists(folder):
        os.makedirs(folder)

    for data_object in data_objects:
        response = requests.get(f"https://api.powercubus.com.br/v1/{data_object}/list?timetable_id={project_id}",
                                headers=my_headers)

        f = open(f"{folder}/{data_object}.json", "w")
        f.write(json.dumps(response.json(), indent=4, sort_keys=True))
        f.close()


if __name__ == '__main__':
    main()
