from openpyxl import load_workbook
import win32com.client as win32
from pathlib import Path
import pandas as pd
import shutil
from pathlib import Path
from datetime import datetime

# Caminho do arquivo Excel com a macro
CAMINHO_EXCEL_ATUALIZAR = "S:/OBSERVATORIO/PROJETOS/BI e Afins/Opera facil/Atualizar - OPERA FACIL.xlsm"
CAMINHO_EXCEL_APRESENTACAO = "S:/OBSERVATORIO/PROJETOS/BI e Afins/Opera facil/apresentação - OPERA FACIL.xlsm"

def backup_relatorio(competencia):
    caminho = "S:/OBSERVATORIO/PROJETOS/BI e Afins/Opera facil/analise-opera-facil/relatorios_simplificados"
    competencia = competencia.replace("/", "-")

    pasta_backup = Path(caminho) / "historico" / competencia
    pasta_backup.mkdir(parents=True, exist_ok=True)

    try:
        shutil.copy2("relatorios_simplificados/separarNeomater_SIMPLIFICADO.xlsx", pasta_backup / "separarNeomater_SIMPLIFICADO.xlsx")
        shutil.copy2("relatorios_simplificados/separarNeotin_SIMPLIFICADO.xlsx", pasta_backup / "separarNeotin_SIMPLIFICADO.xlsx")
        shutil.copy2("relatorios_simplificados/separarPediatrico_SIMPLIFICADO.xlsx", pasta_backup / "separarPediatrico_SIMPLIFICADO.xlsx")
    except Exception as e:
        print(f"❌ Erro ao criar backup dos relatórios: {e}")
        pass
    finally:
        pass

def backup_arquivo():
    
    timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M")
    
    caminhoBackup = "S:/OBSERVATORIO/PROJETOS/BI e Afins/Opera facil/backup"

    backup_atualizar = "{}/Backup_Atualizar_OPERA_FACIL_".format(caminhoBackup) + timestamp + ".xlsm"

    backup_apresentacao = "{}/Backup_Apresentacao_OPERA_FACIL_".format(caminhoBackup) + timestamp + ".xlsm"
    try:

        shutil.copy2(CAMINHO_EXCEL_ATUALIZAR, backup_atualizar)
        shutil.copy2(CAMINHO_EXCEL_APRESENTACAO, backup_apresentacao)
        
        return backup_atualizar, backup_apresentacao
        
    except Exception as e:
        print(f"❌ Erro ao criar backup: {e}")
        return None, None

def executar_macro_atualizar():
    macros = ["juntarNeomater","juntarNeotin", "juntarProntobaby"]
    
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