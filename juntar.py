import win32com.client as win32

# Caminho do arquivo Excel com a macro
CAMINHO_EXCEL = r"C:\Users\luana\Downloads\exel1.xlsm"

# Nome da macro
NOME_MACRO = "juntar"

# Iniciar Excel em segundo plano
excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

# Abrir o arquivo
wb = excel.Workbooks.Open(CAMINHO_EXCEL)

# Executar a macro
excel.Application.Run(NOME_MACRO)

# Salvar e fechar
wb.Close(SaveChanges=True)
excel.Quit()

# Limpeza (boa prática)
del wb
del excel