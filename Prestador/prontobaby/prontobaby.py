import pandas as pd
import os

from procedimentos import (
    pacote_otorrino,
    pacote_geral,
    pacote_oftalmo,
    pacote_hispospadia,
    pacote_inguinal,
    pacote_hidrocele,
    pacote_adeno,
    pacote_amig,
    pacote_amig_adeno,
    pacote_estrabismo,
    pacote_nasal,
    pacote_orqui,
    pacote_plastica,
    pacote_postec,
    pacote_umbilical
)

def analisar_prontobaby():
    arquivo = "relatorios_simplificados/separarPediatrico_SIMPLIFICADO.xlsx"
    municipios = ["RJ - Belford Roxo", "RJ - Duque de Caxias", "RJ - Itaguaí", "RJ - Japeri", 
                  "RJ - Magé", "RJ - Mesquita", "RJ - Nilópolis", "RJ - Nova Iguaçu", 
                  "RJ - Paracambi", "RJ - Queimados", "RJ - Seropédica", "RJ - São João de Meriti"]
    
    # Criar diretório de resultados se não existir
    os.makedirs("Prestador/prontobaby/resultado", exist_ok=True)
    
    # Lista para armazenar os dados consolidados
    dados_consolidados = []
    nao_listados_consolidados = []
    
    # Primeiro processar todos os municípios e coletar os dados
    for municipio in municipios:
        try:
            print(f"\n=== PROCESSANDO {municipio} ===")
            
            # Inicializar todos os pacotes
            otorrino = pacote_otorrino()
            geral = pacote_geral()
            oftalmo = pacote_oftalmo()
            hispospadia = pacote_hispospadia()
            inguinal = pacote_inguinal()
            hidrocele = pacote_hidrocele()
            adeno = pacote_adeno()
            amig = pacote_amig()
            amig_adeno = pacote_amig_adeno()
            estrabismo = pacote_estrabismo()
            nasal = pacote_nasal()
            orqui = pacote_orqui()
            plastica = pacote_plastica()
            postec = pacote_postec()
            umbilical = pacote_umbilical()

            # Ler a tabela
            tabela = pd.read_excel(arquivo)
            
            coluna_procedimento = municipio
            coluna_quantidade = "Quantidade {}".format(municipio)
            
            # Verificar se as colunas existem na tabela
            if coluna_procedimento not in tabela.columns or coluna_quantidade not in tabela.columns:
                print(f"  Aviso: Colunas para '{municipio}' não encontradas na tabela!")
                # Adicionar dados vazios para este município
                dados_municipio = {
                    "Procedimento": [],
                    municipio: []
                }
                dados_consolidados.append(dados_municipio)
                continue
            
            # Definir grupos de procedimentos
            grupos = {
                "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OTORRINO": otorrino or [],
                "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO CIRURGIA GERAL": geral or [],
                "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OFTALMOLOGISTA": oftalmo or [],
                "ADENOIDECTOMIA PEDIÁTRICO": adeno or [],
                "AMIGDALECTOMIA - PEDIATRICO": amig or [],
                "AMIGDALECTOMIA COM ADENOIDECTOMIA - PEDIATRICO": amig_adeno or [],
                "TRATAMENTO CIRÚRGICO DE PERFURAÇÃO DO SEPTO NASAL - PEDIATRICO": nasal or [],
                "CORREÇÃO CIRÚRGICA DE ESTRABISMO (ACIMA DE 2 MUSCULOS) - PEDIATRICO": estrabismo or [],
                "HERNIOPLASTIA INGUINAL (BILATERAL) - PEDIATRICO": inguinal or [],
                "HERNIOPLASTIA UMBILICAL - PEDIATRICO": umbilical or [],
                "ORQUIDOPEXIA BILATERAL - PEDIATRICO": orqui or [],
                "TRATAMENTO CIRÚRGICO DE HIDROCELE - PEDIATRICO": hidrocele or [],
                "CORRECAO DE HIPOSPADIA (1º TEMPO) - PEDIATRICO": hispospadia or [],
                "PLASTICA TOTAL DO PENIS - PEDIATRICO": plastica or [],
                "POSTECTOMIA - PEDIATRICO": postec or [],
            }
            
            # Processar todos os grupos
            totais = {}
            for nome_grupo, lista_procedimentos in grupos.items():
                total_grupo = 0
                for procedimento in lista_procedimentos:
                    if procedimento is None:
                        continue
                    
                    mask = tabela[coluna_procedimento].astype(str) == str(procedimento)
                    quantidade = tabela.loc[mask, coluna_quantidade].sum()
                    
                    try:
                        quantidade_num = float(quantidade) if not pd.isna(quantidade) else 0
                    except (ValueError, TypeError):
                        quantidade_num = 0
                    
                    total_grupo += quantidade_num
                
                totais[nome_grupo] = total_grupo
            
            # Adicionar dados deste município à lista consolidada
            dados_municipio = {
                "Procedimento": list(totais.keys()),
                municipio: list(totais.values())
            }
            dados_consolidados.append(dados_municipio)
            
            # Imprimir resultados no console
            print(f"Resultados para {municipio}:")
            soma_total = sum(totais.values())
            for nome, total in totais.items():
                print(f"  {nome}: {total}")
            print(f"  Soma Total: {soma_total}")
            
            # Identificar procedimentos não mapeados (SOMENTE SE HOUVER PROCEDIMENTOS NA COLUNA)
            procedimentos_na_coluna = tabela[coluna_procedimento].dropna().unique()
            if len(procedimentos_na_coluna) > 0:
                todos_procedimentos_conhecidos = []
                for lista in grupos.values():
                    for p in lista:
                        if p and str(p).strip():
                            todos_procedimentos_conhecidos.append(str(p))
                
                todos_procedimentos_conhecidos = list(set(todos_procedimentos_conhecidos))
                
                # Obter procedimentos da coluna do município
                procedimentos_na_coluna = [str(p).strip() for p in procedimentos_na_coluna if p and str(p).strip()]
                
                # Identificar procedimentos não mapeados
                procedimentos_nao_mapeados = {}
                for procedimento in procedimentos_na_coluna:
                    if procedimento not in todos_procedimentos_conhecidos:
                        mask = tabela[coluna_procedimento].astype(str) == procedimento
                        quantidade = tabela.loc[mask, coluna_quantidade].sum()
                        
                        try:
                            quantidade_num = float(quantidade) if not pd.isna(quantidade) else 0
                        except (ValueError, TypeError):
                            quantidade_num = 0
                        
                        if quantidade_num > 0:  # Só incluir se tiver quantidade > 0
                            procedimentos_nao_mapeados[procedimento] = quantidade_num
                
                # Adicionar não listados à lista consolidada APENAS SE EXISTIREM
                if procedimentos_nao_mapeados:
                    procedimentos_nao_mapeados = dict(sorted(
                        procedimentos_nao_mapeados.items(), 
                        key=lambda x: x[1], 
                        reverse=True
                    ))
                    
                    print(f"\n  Procedimentos não mapeados encontrados: {len(procedimentos_nao_mapeados)}")
                    total_nao_mapeado = sum(procedimentos_nao_mapeados.values())
                    print(f"    Total não mapeado: {total_nao_mapeado}")
                    
                    # Adicionar dados não listados à lista consolidada
                    nao_listados_municipio = {
                        "Procedimento Nao listados": list(procedimentos_nao_mapeados.keys()),
                        municipio: list(procedimentos_nao_mapeados.values()),
                    }
                    nao_listados_consolidados.append(nao_listados_municipio)
                else:
                    print(f"  Nenhum procedimento não mapeado encontrado para {municipio}")
            else:
                print(f"  Nenhum procedimento encontrado na coluna para {municipio}")
            
        except FileNotFoundError:
            print(f"  ERRO: Arquivo {arquivo} não encontrado!")
            
        except Exception as e:
            print(f"  ERRO em {municipio}: {str(e)}")
    
    # AGORA CRIAR OS ARQUIVOS CONSOLIDADOS
    print(f"\n=== CRIANDO ARQUIVOS CONSOLIDADOS ===")
    
    # 1. Criar DataFrame consolidado com todos os municípios lado a lado
    if dados_consolidados:
        # Começar com o primeiro município
        df_consolidado = pd.DataFrame(dados_consolidados[0])
        
        # Juntar os outros municípios
        for i in range(1, len(dados_consolidados)):
            if dados_consolidados[i]["Procedimento"]:
                df_temp = pd.DataFrame(dados_consolidados[i])
                # Mesclar pelo nome do procedimento
                df_consolidado = pd.merge(df_consolidado, df_temp, on="Procedimento", how="outer")
            else:
                # Se não houver dados, adicionar coluna vazia
                municipio_nome = list(dados_consolidados[i].keys())[1] if len(dados_consolidados[i]) > 1 else "Município Desconhecido"
                df_consolidado[municipio_nome] = 0
        
        # Ordenar os procedimentos
        ordem_procedimentos = list(dados_consolidados[0]["Procedimento"]) if dados_consolidados[0]["Procedimento"] else []
        df_consolidado = df_consolidado.set_index("Procedimento").loc[ordem_procedimentos].reset_index()
        
        # Preencher NaN com 0
        df_consolidado = df_consolidado.fillna(0)
        
        # Salvar arquivo consolidado
        df_consolidado.to_excel("Prestador/prontobaby/resultado/relatorio-prontobaby.xlsx", index=False)
        print("✓ Arquivo consolidado criado: relatorio-prontobaby.xlsx")
    
    # 2. Criar DataFrame consolidado para procedimentos não listadosa
    if nao_listados_consolidados:
        # Começar com o primeiro município que tem não listados
        df_nao_listados = pd.DataFrame(nao_listados_consolidados[0])
        
        # Juntar os outros municípios
        for i in range(1, len(nao_listados_consolidados)):
            if nao_listados_consolidados[i]["Procedimento Nao listados"]:
                df_temp = pd.DataFrame(nao_listados_consolidados[i])
                # Mesclar pelo nome do procedimento
                df_nao_listados = pd.merge(df_nao_listados, df_temp, on="Procedimento Nao listados", how="outer")
        
        # Preencher NaN com 0
        df_nao_listados = df_nao_listados.fillna(0)
        
        # Salvar arquivo consolidado de não listados
        df_nao_listados.to_excel("Prestador/prontobaby/resultado/CONSOLIDADO_NAO_LISTADOS.xlsx", index=False)
        print("✓ Arquivo consolidado de não listados criado: CONSOLIDADO_NAO_LISTADOS.xlsx")
    else:
        print("✗ Nenhum procedimento não listado encontrado para criar arquivo consolidado")
    
    print(f"\n=== PROCESSAMENTO CONCLUÍDO ===")
    print(f"Total de municípios processados: {len(municipios)}")

# Executar a função
if __name__ == "__main__":
    analisar_prontobaby()