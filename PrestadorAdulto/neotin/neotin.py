import pandas as pd
import os
from procedimentos import procedimentos  # <-- agora você importa só isso


def carregar(tabela, nome_coluna):
    return [
        x for x in tabela[nome_coluna].to_list()
        if pd.notna(x) and x not in [None, '']
    ]


def analisar_neotin():
    arquivo = "relatorios_simplificados/separarneotin_SIMPLIFICADO.xlsx"
    municipios = [
        "RJ - Belford Roxo", "RJ - Duque de Caxias", "RJ - Itaguaí", "RJ - Japeri",
        "RJ - Magé", "RJ - Mesquita", "RJ - Nilópolis", "RJ - Nova Iguaçu",
        "RJ - Paracambi", "RJ - Queimados", "RJ - Seropédica", "RJ - São João de Meriti"
    ]

    os.makedirs("PrestadorAdulto/neotin/resultado", exist_ok=True)

    # 🔥 CARREGA DB UMA VEZ SÓ
    tabela_base = pd.read_excel("db.xlsx")

    # 🔥 GERA TODAS AS LISTAS AUTOMATICAMENTE
    listas = {
        nome: carregar(tabela_base, coluna)
        for nome, coluna in procedimentos.items()
    }

    # 🔥 NOMES NA ORDEM (igual Excel)
    nomes_grupos = list(procedimentos.values())

    dados_consolidados = {
        nome: {municipio: 0 for municipio in municipios}
        for nome in nomes_grupos
    }

    nao_listados_consolidados = {}

    for municipio in municipios:
        try:
            print(f"\n=== PROCESSANDO {municipio} ===")

            tabela = pd.read_excel(arquivo)

            coluna_proc = municipio
            coluna_qtd = f"Quantidade {municipio}"

            if coluna_proc not in tabela.columns or coluna_qtd not in tabela.columns:
                print(f"  Aviso: Colunas não encontradas")
                continue

            # 🔥 GRUPOS AUTOMÁTICOS
            grupos = {
                procedimentos[chave]: listas[chave]
                for chave in procedimentos
            }

            soma_total = 0

            for nome_grupo, lista_proc in grupos.items():
                total = 0

                for proc in lista_proc:
                    mask = tabela[coluna_proc].astype(str) == str(proc)
                    qtd = tabela.loc[mask, coluna_qtd].sum()

                    try:
                        qtd = float(qtd) if not pd.isna(qtd) else 0
                    except:
                        qtd = 0

                    total += qtd

                dados_consolidados[nome_grupo][municipio] = total
                soma_total += total

            print(f"Soma Total: {soma_total}")

        except Exception as e:
            print(f"Erro em {municipio}: {e}")

    # 🔥 GERAR DATAFRAME FINAL
    df = pd.DataFrame.from_dict(dados_consolidados, orient='index')
    df = df[municipios]
    df = df.loc[nomes_grupos]

    df.loc['TOTAL'] = df.sum()

    # RELATÓRIO COMPLETO
    caminho = "PrestadorAdulto/neotin/resultado/relatorio_final.xlsx"
    df.reset_index().rename(columns={"index": "Procedimento"}).to_excel(caminho, index=False)

    print(f"\n✅ Relatório gerado: {caminho}")


if __name__ == "__main__":
    analisar_neotin()