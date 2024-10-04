"""Script para importação de disponibilidades de aulas flutuantes

Script para importação de disponibilidades de aulas flutuantes para bloquear dias de
semana para essas aulas no PowerCubus.
Recebe como entrada o id do projeto e o token de acesso ao sistema e utiliza o código externo da aula para descobrir se
deve bloquear os dias de semana ou bloquear o domingo quando não é flutuante.
Acessa a API do PowerCubus para realizar as importações.

Pode ser utilizado pelo analista de TI para auxiliar no processo de elaboração de horários no PowerCubus.

Desenvolvido por Vinicius Tozo
Última atualização: 12/11/2021
"""
import json

import requests


def get_events(project_id, access_token):
    my_headers = {'Authorization': f'Bearer {access_token}'}
    events = requests.get(f'https://api.powercubus.com.br/v1/events/list?timetable_id={project_id}',
                          headers=my_headers)

    return events.json()


def update_availability(events_json, project_id, access_token):
    my_headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-type': 'application/json',
        'Accept': 'application/json'
    }

    # Cria um dicionário com o nome das disciplinas
    dict_subjects = {}
    subjects = requests.get(f"https://api.powercubus.com.br/v1/subjects/list?timetable_id={project_id}",
                            headers=my_headers)
    for subject in subjects.json():
        dict_subjects[subject['id']] = subject['name']

    # Para cada aula
    for e in events_json:

        if ';' not in e['external_id']:
            flutuante = False
        else:
            flutuante = e['external_id'].split(';')[4] == '1'

        subject_name = dict_subjects.get(e['subject_id'], '').upper()
        if 'PROJETO FINAL' in subject_name or \
                'ESTÁGIO' in subject_name or \
                'PROJETO INTEGRADOR' in subject_name:
            flutuante = True

        availability = requests.get(f'https://api.powercubus.com.br/v1/events/list-availability?'
                                    f'timetable_id={project_id}&event_id={e["id"]}', headers=my_headers).json()

        # Modifica apenas quando necessário
        modified = False
        for i in range(len(availability)):
            if flutuante and availability[i]['day_id'] == 1 and availability[i]['unavailable'] != 0:
                availability[i]['unavailable'] = 0
                modified = True
            elif flutuante and availability[i]['day_id'] != 1 and availability[i]['unavailable'] != 2:
                availability[i]['unavailable'] = 2
                modified = True
            elif not flutuante and availability[i]['day_id'] == 1 and availability[i]['unavailable'] != 2:
                availability[i]['unavailable'] = 2
                modified = True
            elif not flutuante and availability[i]['day_id'] != 1 and availability[i]['unavailable'] != 0:
                availability[i]['unavailable'] = 0
                modified = True

        if not modified:
            print(f'{e["id"]: <100} Disponibilidade já estava atualizada no PowerCubus')
            continue

        data = []
        for a in availability:
            data.append({
                "day_id": int(a['day_id']),
                "period_id": int(a['period']['id']),
                "unavailable": int(a['unavailable']),
            })

        data = json.dumps(data)

        response = requests.put(f'https://api.powercubus.com.br/v1/events/update-availability?'
                                f'timetable_id={project_id}&event_id={e["id"]}', headers=my_headers, data=data)
        print(f'{e["id"]: <100} Atualização da disponibilidade retornou código '
              f'{response.status_code} ({response.reason})')


def main():
    project_id = input('ID do Projeto: ')
    access_token = input('Token de acesso: ')

    events_json = get_events(project_id, access_token)

    update_availability(events_json, project_id, access_token)


if __name__ == "__main__":
    main()
