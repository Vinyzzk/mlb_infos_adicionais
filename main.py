import requests
import os
import pandas as pd
import openpyxl


def main():
    data = []
    
    try:
        with open("mlbs.txt", "r") as mlbs:
            mlbs = mlbs.read().splitlines()

            print("Gerando excel...")

            for mlb in mlbs:
                url = requests.get(f"https://api.mercadolibre.com/items/{mlb}")
                response = url.json()
                created_at = response["date_created"][:10]

                created_at = created_at.split("-")

                year = created_at[0]
                month = created_at[1]
                day = created_at[2]

                created_at = f"{day}/{month}/{year}"

                sold_quantity = response["sold_quantity"]
                catalog = response["catalog_listing"]
                    
                data.append({'MLB': mlb, 'Criado em': created_at, 'Vendas': sold_quantity, 'Catalogo': catalog})

            df = pd.DataFrame(data)
            df.to_excel("result.xlsx", index=False, engine="openpyxl")
    except FileNotFoundError:
        try:
            df = pd.read_excel("mlbs.xlsx")
            column = df["MLB"]
            mlbs = column.values
            
            for mlb in mlbs:
                    url = requests.get(f"https://api.mercadolibre.com/items/{mlb}")
                    response = url.json()
                    created_at = response["date_created"][:10]

                    created_at = created_at.split("-")

                    year = created_at[0]
                    month = created_at[1]
                    day = created_at[2]

                    created_at = f"{day}/{month}/{year}"

                    sold_quantity = response["sold_quantity"]
                    catalog = response["catalog_listing"]
                    
                    data.append({'MLB': mlb, 'Criado em': created_at, 'Vendas': sold_quantity, 'Catalogo': catalog})

            df = pd.DataFrame(data)
            df.to_excel("result.xlsx", index=False, engine="openpyxl")
        except FileNotFoundError:
            print("[!] Arquivo \"mlbs.txt\" ou \"mlbs.xlsx\" nao encontrado")
            input("[+] Pressione ENTER para finalizar")


if __name__ == "__main__":
    main()
