import win32com.client as win32

# Caminho do arquivo Excel com a macro
CAMINHO_EXCEL = r"c:\Users\dalves\Downloads\Atualizar - OPERA FACIL.xlsm"

# Nome da macro
def executar_macro():
    macros = ["juntarNeomater","juntarNeotin", "juntarProntobaby"]
# Iniciar Excel em segundo plano
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Abrir o arquivo
    wb = excel.Workbooks.Open(CAMINHO_EXCEL)

    # Executar a macro
    for macro in macros:
        excel.Application.Run(macro)

    # Salvar e fechar
    wb.Close(SaveChanges=True)
    excel.Quit()

    # Limpeza (boa prática)
    del wb
    del excel