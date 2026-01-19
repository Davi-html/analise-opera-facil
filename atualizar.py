import win32com.client as win32
from pathlib import Path

# Caminho do arquivo Excel com a macro
CAMINHO_EXCEL_ATUALIZAR = Path.home() / "Downloads" / "Atualizar - OPERA FACIL.xlsm"
CAMINHO_EXCEL_APRESENTACAO = Path.home() / "Downloads" / "apresentação - OPERA FACIL.xlsm"

# Nome da macro
def executar_macro_atualizar():
    macros = ["juntarNeomater","juntarNeotin", "juntarProntobaby"]
# Iniciar Excel em segundo plano
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Abrir o arquivo
    wb = excel.Workbooks.Open(CAMINHO_EXCEL_ATUALIZAR)

    # Executar a macro
    for macro in macros:
        excel.Application.Run(macro)

    # Salvar e fechar
    wb.Close(SaveChanges=True)
    excel.Quit()

    # Limpeza (boa prática)
    del wb
    del excel

def executar_macro_apresentacao():
    macro = "ConsolidarTodosDados"
    # Iniciar Excel em segundo plano
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Abrir o arquivo
    wb = excel.Workbooks.Open(CAMINHO_EXCEL_APRESENTACAO)

    # Executar a macro
    excel.Application.Run(macro)

    # Salvar e fechar
    wb.Close(SaveChanges=True)
    excel.Quit()

    # Limpeza (boa prática)
    del wb
    del excel
