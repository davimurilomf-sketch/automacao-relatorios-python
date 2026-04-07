import xlwings as xw
import time
from datetime import timedelta

# CONFIGURAÇÃO DE CADA RELATÓRIO
arquivos = [
    {
        "caminho": r"X:\Hora a Hora\HORA HORA CLIENTE (VIA+CSED+IPA).xlsx",
        "aba": ["HORAxHORA Via Varejo","HORAxHORA Cruzeiro","HORAxHORA Ipanema"]
    },
    {
        "caminho": r"X:\Hora a Hora\HORA X HORA COLCHÃO GRUPO B.xlsx",
        "aba": "Base"
    },
    {
        "caminho": r"X:\Hora a Hora\HORA X HORA COLCHÃO RECOVERY.xlsx",
        "aba": "BASE"

    },
    {
        "caminho": r"X:\Hora a Hora\HORA x HORA OPERACIONAL (VIA+CSED+IPA).xlsb",
        "aba": ["COMPARATIVO DIA","COMPARATIVO HORA"]

    }
]

app = xw.App(visible=True)

try:
    for item in arquivos:
        print(f"Processando: {item['caminho']}")

        wb = app.books.open(item["caminho"])

        try:
            # atualizar tudo direto
            wb.api.RefreshAll()
            app.api.CalculateUntilAsyncQueriesDone()

            # esperar terminar
            while app.api.CalculationState != 0:
                time.sleep(1)

            # atualizar tabela dinâmica
            for sheet in wb.sheets:
                for pivot in sheet.api.PivotTables():
                    pivot.RefreshTable()

            wb.save()
            print("Arquivo atualizado com sucesso!")

        except Exception as e:
            print(f"Erro nesse arquivo: {e}")

        finally:
            wb.close()

finally:
    app.quit()