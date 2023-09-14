import requests
import os
import pandas as pd
import openpyxl


def main():
    data = []
    with open("mlbs.txt", "r") as mlbs:
        mlbs = mlbs.read().splitlines()

        print("Gerando excel...")

        for mlb in mlbs:
            url = requests.get(f"https://api.mercadolibre.com/items/{mlb}")
            response = url.json()
            created_at = response["date_created"][:10]
            sold_quantity = response["sold_quantity"]
            data.append({'MLB': mlb, 'Criado em': created_at, 'Vendas': sold_quantity})

        df = pd.DataFrame(data)
        df.to_excel("result.xlsx", index=False, engine="openpyxl")


if __name__ == "__main__":
    main()