import xlwings as xw
import time
from datetime import timedelta

# CONFIGURAÇÃO DE CADA RELATÓRIO
arquivos = [
    {
        "caminho": r"X:Indicadores operacionais\Relatório de Preventivo - Cruzeiro.xlsx",
        "aba": "RESUMO"
    },
    {
        "caminho": r"X:Indicadores operacionais\Relatório de Preventivo - Ipanema.xlsx",
        "aba": "RESUMO"
    },
    {
        "caminho": r"X:Indicadores operacionais\Relatório de Preventivo - Via Varejo.xlsx",
        "aba": "RESUMO"

    },
    {
        "caminho": r"X:Indicadores operacionais\SUSP - VIA VAREJO.xlsb",
        "aba": "RESUMO"

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