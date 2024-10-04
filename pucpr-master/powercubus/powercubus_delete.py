"""Script para APAGAR os dados de um projeto no PowerCubus

Script para APAGAR todos os dados de um projeto no sistema via API.
Recebe como entrada o id do projeto e o token de acesso ao sistema e apaga os dados utilizando a API.
As listas 'data_objects' e 'data_objects_singular' contêm os dados que devem ser apagados, exemplo:
    data_objects = ['events']                       # para apagar apenas aulas
    data_objects_singular = ['event']               # para apagar apenas aulas
    data_objects = ['events', 'teachers']           # para apagar aulas e professores
    data_objects_singular = ['event', 'teacher']    # para apagar aulas e professores
Observação: os itens presentes nas duas listas devem ser correspondentes, um no plural e um no singular, na mesma ordem.

Pode ser utilizado pelo analista de TI para auxiliar no processo de elaboração de horários no PowerCubus.

Desenvolvido por Vinicius Tozo
Última atualização: 03/02/2022
"""
import requests


def main():
    project_id = input('ID do Projeto: ')
    access_token = input('Token de acesso: ')

    delete_all(project_id, access_token)


def delete_all(project_id, access_token):
    my_headers = {'Authorization': f'Bearer {access_token}'}

    data_objects = ['events']
    data_objects_singular = ['event']

    for data_object in data_objects:
        response = requests.get(f"https://api.powercubus.com.br/v1/{data_object}/list?timetable_id={project_id}",
                                headers=my_headers)

        if response.status_code != 200:
            print(data_object, response.status_code)
            continue

        for item in response.json():
            print(item['id'], end='\t')
            response = requests.delete(
                f"https://api.powercubus.com.br/v1/{data_object}/remove?timetable_id={project_id}&"
                f"{data_objects_singular[data_objects.index(data_object)]}_id={item['id']}",
                headers=my_headers)
            print(response.request.url, response.status_code, response.reason)


if __name__ == '__main__':
    main()
