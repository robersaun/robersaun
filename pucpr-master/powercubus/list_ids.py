"""Script para listar IDs de aulas do PowerCubus

Script para listagem de IDs a partir de um arquivo json contendo a exportação das aulas (events) do sistema.
Recebe como entrada o arquivo json e imprime em tela a lista de IDs.

Pode ser utilizado pelo analista de TI para auxiliar no processo de elaboração de horários no PowerCubus.

Desenvolvido por Vinicius Tozo
Última atualização: 07/01/2022
"""
import json
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def main():
    Tk().withdraw()
    file_name_json = askopenfilename(filetypes=[('json', '.json')])

    # Opening JSON file
    f = open(file_name_json)

    # returns JSON object as
    # a dictionary
    data = json.load(f)

    # Iterating through the json
    # list
    for i in data:
        print(str(i['id']) + '\t' + i['external_id'])


if __name__ == '__main__':
    main()
