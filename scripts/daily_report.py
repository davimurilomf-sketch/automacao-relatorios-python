try:
    import xlwings as xw
except ImportError as e:
    raise ImportError(
        "O módulo xlwings não está instalado. Instale com: pip install xlwings"
    ) from e
import time
from datetime import timedelta

# CONFIGURAÇÃO DE CADA RELATÓRIO
arquivos = [
    {
        "caminho": r"data/relatorio.xlsx",
        "aba": "Pagina",
        "celula": "I8"
    },
    {
        "caminho": r"data/relatorio.xlsx",
        "aba": "Pagina",
        "celula": "H5"
    },
    {
        "caminho": r"data/relatorio.xlsx",
        "aba": "Pagina",
        "celula": "H5"
    }
]

app = xw.App(visible=False)

try:
    for item in arquivos:
        print(f"Processando: {item['caminho']}")

        wb = app.books.open(item["caminho"])

        try:
            ws = wb.sheets[item["aba"]]

            # pegar data
            celula = ws.range(item["celula"])
            data_atual = celula.value

            # aumentar 1 dia
            nova_data = data_atual + timedelta(days=1)
            celula.value = nova_data

            # atualizar tudo
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
            print("OK!")

        except Exception as e:
            print(f"Erro nesse arquivo: {e}")

        finally:
            wb.close()

finally:
    app.quit()
