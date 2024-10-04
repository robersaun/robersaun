import pandas


def main():
    # Lê o arquivo
    df = pandas.read_excel("Base_para_Python_Estágio.xlsx")

    # Seleciona as vendas da marca Maçã
    df_maca = df.loc[df.Marca == "Maçã"]

    # Imprime a soma dos valores
    print(f"Soma dos valores da marca Maçã: {df_maca.Valor.sum()}")

    # Seleciona as vendas da marca Banana
    df_banana = df.loc[df.Marca == "Banana"]

    # Salva o arquivo
    df_banana.to_excel("vendas_banana.xlsx")


if __name__ == '__main__':
    main()
