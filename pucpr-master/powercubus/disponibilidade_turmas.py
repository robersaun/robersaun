"""Script para importação de disponibilidades de turmas

Script para importação de disponibilidades de turmas conforme o turno indicado no nome da turma.
Recebe como entrada o id do projeto e o token de acesso ao sistema e utiliza o nome da turma para descobrir quais
turnos deve bloquear.
Acessa a API do PowerCubus para realizar as importações.

Pode ser utilizado pelo analista de TI para auxiliar no processo de elaboração de horários no PowerCubus.

Desenvolvido por Vinicius Tozo
Última atualização: 12/11/2021
"""
import json

import requests


def get_classes(project_id, access_token):
    my_headers = {'Authorization': f'Bearer {access_token}'}
    classes = requests.get(f'https://api.powercubus.com.br/v1/classes/list?timetable_id={project_id}',
                           headers=my_headers)

    return classes.json()


def update_availability(classes_json, project_id, access_token):
    my_headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-type': 'application/json',
        'Accept': 'application/json'
    }

    shift_ids = {'M': [1], 'T': [2], 'N': [3], 'I': [1, 2]}

    # Para cada turma
    for c in classes_json:
        shift = c['name'].split(' - ')[3].strip()
        shift_id = shift_ids.get(shift)

        availability = requests.get(f'https://api.powercubus.com.br/v1/classes/list-availability?'
                                    f'timetable_id={project_id}&class_id={c["id"]}', headers=my_headers).json()

        # Verifica se a disponibilidade foi alterada manualmente
        is_default = True
        for i in range(len(availability)):
            if availability[i]['unavailable'] != 0:
                is_default = False

        # Apenas muda a disponibilidade se ela não foi alterada antes
        if not is_default:
            print(f'{c["name"]: <100} Disponibilidade já havia sido modificada')
            continue

        # Modifica apenas quando necessário
        modified = False
        for i in range(len(availability)):
            if availability[i]['period']['shift'] in shift_id and availability[i]['unavailable'] != 0 \
                    and availability[i]['period']['start_time'] != '07:05':
                availability[i]['unavailable'] = 0
                modified = True
            elif availability[i]['period']['shift'] not in shift_id and availability[i]['unavailable'] != 2:
                availability[i]['unavailable'] = 2
                modified = True

            # Caso de exceção - Horário das 07:05 não tem aula mas deve ficar no sistema
            # para importação de disponibilidade de professores
            if availability[i]['period']['start_time'] == '07:05' and availability[i]['unavailable'] != 2:
                availability[i]['unavailable'] = 2
                modified = True

        if not modified:
            print(f'{c["name"]: <100} Disponibilidade já estava atualizada no PowerCubus')
            continue

        data = []
        for a in availability:
            data.append({
                "day_id": int(a['day_id']),
                "period_id": int(a['period']['id']),
                "unavailable": int(a['unavailable']),
            })

        data = json.dumps(data)

        response = requests.put(f'https://api.powercubus.com.br/v1/classes/update-availability?'
                                f'timetable_id={project_id}&class_id={c["id"]}', headers=my_headers, data=data)
        print(f'{c["name"]: <100} Atualização da disponibilidade retornou código '
              f'{response.status_code} ({response.reason})')


def main():
    project_id = input('ID do Projeto: ')
    access_token = input('Token de acesso: ')

    classes_json = get_classes(project_id, access_token)

    update_availability(classes_json, project_id, access_token)


if __name__ == "__main__":
    main()
