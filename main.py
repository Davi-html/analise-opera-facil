from PrestadorAdulto.neotin.neotin import analisar_neotin
from separarRelatorio.main import processar_todos_arquivos_simplificado

def main():
    processar_todos_arquivos_simplificado()
    
    analisar_neotin()

main()