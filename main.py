from PrestadorAdulto.neotin.neotin import analisar_prestador
from analise_financeiro.financeiro import analise_financeiro
from separarRelatorio.main import processar_todos_arquivos_simplificado

def main():
    processar_todos_arquivos_simplificado()
    
    prestadores = ["Catarina", "Desam", "neotin", "Pronil", "Uroclin", "Vivermais"]

    for prestador in prestadores:
        analisar_prestador(prestador)


    analise_financeiro("20/03 a 19/04", "2026", prestadores)

main()