"""Script para importação de disponibilidades de aulas

Script para importação de disponibilidades de aulas para replicar horários gerados no aSc Timetables.
Recebe como entrada o mesmo template que foi importado no PowerCubus, o id do projeto e o token de acesso ao sistema.
Acessa a API do PowerCubus para realizar as importações.

Utilizado pelo analista de TI no processo de replicação de horários do aSc no PowerCubus.

Desenvolvido por Vinicius Tozo
Última atualização: 02/02/2022
"""
import json
import time
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas
import requests


def get_dataframe():
    # Seleciona o arquivo do aSc Timetables
    print('Selecione o arquivo excel:')
    Tk().withdraw()
    file_name_xlsx = askopenfilename(filetypes=[('xlsx', '.xlsx')])
    print(f'\tArquivo selecionado: {file_name_xlsx}')

    # Lê o arquivo e grava em um dataframe
    df_lessons = pandas.read_excel(file_name_xlsx)
    df_lessons = df_lessons[['Código de origem da aula', 'Horário Fixo']]

    # Remove linhas que não contém horário fixo
    df_lessons.dropna(subset=['Horário Fixo'], inplace=True)

    return df_lessons


def get_credentials():
    project_id = input('ID do Projeto: ')
    access_token = input('Token de acesso: ')
    return project_id, access_token


def print_dataframe(df_lessons):
    for row, value in df_lessons.iterrows():
        print(f"{{\n\t"
              f"'id': '{value['Código de origem da aula']}'\n\t"
              f"'periods': '{value['Horário Fixo']}'\n"
              f"}},")


def get_events(project_id, access_token):
    my_headers = {'Authorization': f'Bearer {access_token}'}
    events = requests.get(f'https://api.powercubus.com.br/v1/events/list?timetable_id={project_id}',
                          headers=my_headers)

    return events.json()


def print_json(json_text):
    print(json.dumps(json_text, indent=4))


def update_event_availability(project_id, my_headers, event, days, df_lesson_periods):
    time.sleep(2)
    availability = requests.get(f'https://api.powercubus.com.br/v1/events/list-availability?'
                                f'timetable_id={project_id}&event_id={event["id"]}', headers=my_headers).json()

    # Modifica apenas quando necessário
    modified = False

    # Para cada disponibilidade da aula no sistema
    for i in range(len(availability)):

        # Verifica se esta na lista do arquivo excel
        is_available = False
        for df_lesson_period in df_lesson_periods:
            day = days.get(df_lesson_period.split()[0])
            start_time = df_lesson_period.split()[1].zfill(5)

            if availability[i]['day_id'] == day and availability[i]['period']['start_time'] == start_time:
                is_available = True

        # Modifica os valores se necessário
        if is_available and availability[i]['unavailable'] == 2:
            availability[i]['unavailable'] = 0
            modified = True
        elif not is_available and availability[i]['unavailable'] == 0:
            availability[i]['unavailable'] = 2
            modified = True

    # Se não houve modificação passa para a próxima aula
    if not modified:
        print(f'Event {event["external_id"].split(";")[0]: <100} Disponibilidade já estava atualizada no PowerCubus')
        return

    # Se houve atualiza no sistema via API
    data = []
    for a in availability:
        data.append({
            "day_id": int(a['day_id']),
            "period_id": int(a['period']['id']),
            "unavailable": int(a['unavailable']),
        })

    data = json.dumps(data, indent=4)

    response = requests.put(f'https://api.powercubus.com.br/v1/events/update-availability?'
                            f'timetable_id={project_id}&event_id={event["id"]}', headers=my_headers, data=data)
    print(f'Event {event["id"]: <100} Atualização da disponibilidade retornou código '
          f'{response.status_code} ({response.reason})')


def update_class_availability(project_id, my_headers, event, days, df_lesson_periods):
    availability = requests.get(f'https://api.powercubus.com.br/v1/classes/list-availability?'
                                f'timetable_id={project_id}&class_id={event["class_id"]}', headers=my_headers).json()

    # Modifica apenas quando necessário
    modified = False

    # Para cada disponibilidade da aula no sistema
    for i in range(len(availability)):

        # Verifica se esta na lista do arquivo excel
        is_available = False
        for df_lesson_period in df_lesson_periods:
            day = days.get(df_lesson_period.split()[0])
            start_time = df_lesson_period.split()[1].zfill(5)

            if availability[i]['day_id'] == day and availability[i]['period']['start_time'] == start_time:
                is_available = True

        # Modifica os valores se necessário
        if is_available and availability[i]['unavailable'] == 2:
            availability[i]['unavailable'] = 0
            modified = True

    # Se não houve modificação passa para a próxima aula
    if not modified:
        print(f'Class {event["class_id"]: <100} Disponibilidade já estava atualizada no PowerCubus')
        return

    # Se houve atualiza no sistema via API
    data = []
    for a in availability:
        data.append({
            "day_id": int(a['day_id']),
            "period_id": int(a['period']['id']),
            "unavailable": int(a['unavailable']),
        })

    data = json.dumps(data, indent=4)

    response = requests.put(f'https://api.powercubus.com.br/v1/classes/update-availability?'
                            f'timetable_id={project_id}&class_id={event["class_id"]}', headers=my_headers, data=data)
    print(f'Class {event["class_id"]: <100} Atualização da disponibilidade retornou código '
          f'{response.status_code} ({response.reason})')


def update_teacher_availability(project_id, my_headers, event, days, df_lesson_periods):
    availability = requests.get(f'https://api.powercubus.com.br/v1/teachers/list-availability?'
                                f'timetable_id={project_id}&teacher_id={event["teacher_id"]}',
                                headers=my_headers).json()

    # Modifica apenas quando necessário
    modified = False

    # Para cada disponibilidade da aula no sistema
    for i in range(len(availability)):

        # Verifica se esta na lista do arquivo excel
        is_available = False
        for df_lesson_period in df_lesson_periods:
            day = days.get(df_lesson_period.split()[0])
            start_time = df_lesson_period.split()[1].zfill(5)

            if availability[i]['day_id'] == day and availability[i]['period']['start_time'] == start_time:
                is_available = True

        # Modifica os valores se necessário
        if is_available and availability[i]['unavailable'] == 2:
            availability[i]['unavailable'] = 0
            modified = True

    # Se não houve modificação passa para a próxima aula
    if not modified:
        print(f'Teacher {event["teacher_id"]: <100} Disponibilidade já estava atualizada no PowerCubus')
        return

    # Se houve atualiza no sistema via API
    data = []
    for a in availability:
        data.append({
            "day_id": int(a['day_id']),
            "period_id": int(a['period']['id']),
            "unavailable": int(a['unavailable']),
            "unavailable_resource": int(a['unavailable_resource']),
        })

    data = json.dumps(data, indent=4)

    response = requests.put(f'https://api.powercubus.com.br/v1/teachers/update-availability?'
                            f'timetable_id={project_id}&teacher_id={event["teacher_id"]}',
                            headers=my_headers, data=data)
    print(f'Teacher {event["teacher_id"]: <100} Atualização da disponibilidade retornou código '
          f'{response.status_code} ({response.reason})')


def update_availability(events_json, df_lessons, project_id, access_token):
    my_headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-type': 'application/json',
        'Accept': 'application/json'
    }

    days = {'Flutuante': 1, 'Seg': 2, 'Ter': 3, 'Qua': 4, 'Qui': 5, 'Sex': 6, 'Sáb': 7}

    # Para cada aula no PowerCubus
    for event in events_json:
        external_id = event['external_id']

        # Procura a aula no arquivo excel
        df_lesson = df_lessons.loc[df_lessons['Código de origem da aula'] == external_id]

        # Se não encontrar passa para a próxima aula
        if len(df_lesson.index) == 0:
            continue

        df_lesson_periods = df_lesson['Horário Fixo'].iloc[0].split(',')

        update_event_availability(project_id, my_headers, event, days, df_lesson_periods)
        # update_class_availability(project_id, my_headers, event, days, df_lesson_periods)
        # update_teacher_availability(project_id, my_headers, event, days, df_lesson_periods)


def main():
    df_lessons = get_dataframe()
    project_id, access_token = get_credentials()
    events_json = get_events(project_id, access_token)
    update_availability(events_json, df_lessons, project_id, access_token)


if __name__ == "__main__":
    main()
