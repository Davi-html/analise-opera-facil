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
    
    # Dicionário para armazenar os dados consolidados
    dados_consolidados = {}
    nao_listados_consolidados = {}
    
    # Primeiro, inicializar todos os procedimentos com zeros para todos os municípios
    # Definir grupos de procedimentos (apenas os nomes primeiro)
    nomes_grupos = [
        "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OTORRINO",
        "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO CIRURGIA GERAL",
        "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OFTALMOLOGISTA",
        "ADENOIDECTOMIA PEDIÁTRICO",
        "AMIGDALECTOMIA - PEDIATRICO",
        "AMIGDALECTOMIA COM ADENOIDECTOMIA - PEDIATRICO",
        "TRATAMENTO CIRÚRGICO DE PERFURAÇÃO DO SEPTO NASAL - PEDIATRICO",
        "CORREÇÃO CIRÚRGICA DE ESTRABISMO (ACIMA DE 2 MUSCULOS) - PEDIATRICO",
        "HERNIOPLASTIA INGUINAL (BILATERAL) - PEDIATRICO",
        "HERNIOPLASTIA UMBILICAL - PEDIATRICO",
        "ORQUIDOPEXIA BILATERAL - PEDIATRICO",
        "TRATAMENTO CIRÚRGICO DE HIDROCELE - PEDIATRICO",
        "CORRECAO DE HIPOSPADIA (1º TEMPO) - PEDIATRICO",
        "PLASTICA TOTAL DO PENIS - PEDIATRICO",
        "POSTECTOMIA - PEDIATRICO"
    ]
    
    # Inicializar todos os procedimentos com zeros para todos os municípios
    for nome_grupo in nomes_grupos:
        dados_consolidados[nome_grupo] = {municipio: 0 for municipio in municipios}
    
    # Processar todos os municípios
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
                # Manter zeros para este município
                continue
            
            # Definir mapeamento de grupos para listas de procedimentos
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
            
            # Processar todos os grupos para este município
            soma_total = 0
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
                
                # Atualizar valor para este município
                dados_consolidados[nome_grupo][municipio] = total_grupo
                soma_total += total_grupo
            
            # Imprimir resultados no console
            print(f"Resultados para {municipio}:")
            for nome_grupo in nomes_grupos:
                valor = dados_consolidados[nome_grupo][municipio]
                if valor > 0:  # Mostrar apenas os que têm valor > 0
                    print(f"  {nome_grupo}: {valor}")
            
            print(f"  Soma Total: {soma_total}")
            
            # Identificar procedimentos não mapeados
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
                for procedimento in procedimentos_na_coluna:
                    if procedimento not in todos_procedimentos_conhecidos:
                        mask = tabela[coluna_procedimento].astype(str) == procedimento
                        quantidade = tabela.loc[mask, coluna_quantidade].sum()
                        
                        try:
                            quantidade_num = float(quantidade) if not pd.isna(quantidade) else 0
                        except (ValueError, TypeError):
                            quantidade_num = 0
                        
                        if quantidade_num > 0:  # Só incluir se tiver quantidade > 0
                            # Inicializar a entrada para este procedimento se não existir
                            if procedimento not in nao_listados_consolidados:
                                nao_listados_consolidados[procedimento] = {m: 0 for m in municipios}
                            
                            # Adicionar valor para este município
                            nao_listados_consolidados[procedimento][municipio] = quantidade_num
                
                if nao_listados_consolidados:
                    total_nao_mapeado = 0
                    for proc in nao_listados_consolidados:
                        if municipio in nao_listados_consolidados[proc]:
                            total_nao_mapeado += nao_listados_consolidados[proc][municipio]
                    
                    print(f"  Procedimentos não mapeados encontrados: {len([p for p in nao_listados_consolidados if nao_listados_consolidados[p][municipio] > 0])}")
                    print(f"    Total não mapeado: {total_nao_mapeado}")
            
        except FileNotFoundError:
            print(f"  ERRO: Arquivo {arquivo} não encontrado!")
        except Exception as e:
            print(f"  ERRO em {municipio}: {str(e)}")
    
    # CRIAR OS ARQUIVOS CONSOLIDADOS
    print(f"\n=== CRIANDO ARQUIVOS CONSOLIDADOS ===")
    
    # 1. Criar DataFrame consolidado com todos os municípios lado a lado
    if dados_consolidados:
        # Converter dicionário para DataFrame
        df_consolidado = pd.DataFrame.from_dict(dados_consolidados, orient='index')
        
        # Garantir que todas as colunas (municípios) existam na ordem correta
        for municipio in municipios:
            if municipio not in df_consolidado.columns:
                df_consolidado[municipio] = 0
        
        # Reordenar as colunas na ordem dos municípios
        df_consolidado = df_consolidado[municipios]
        
        # Reordenar as linhas (procedimentos) na ordem desejada
        df_consolidado = df_consolidado.loc[nomes_grupos]
        
        # Adicionar uma linha com o TOTAL por município
        totais_municipios = df_consolidado.sum()
        df_consolidado.loc['TOTAL'] = totais_municipios
        
        # ========== RELATÓRIO 1: APENAS VALORES (SEM CABEÇALHOS) ==========
        print("\n--- CRIANDO RELATÓRIO 1: APENAS VALORES ---")
        
        # Criar uma cópia dos dados apenas com valores (sem índices)
        df_apenas_valores = df_consolidado.copy()
        
        # Salvar apenas os valores (sem cabeçalhos de coluna e sem índice)
        caminho_apenas_valores = "Prestador/prontobaby/resultado/relatorio_prontobaby_APENAS_VALORES.xlsx"
        
        # Criar um ExcelWriter para salvar sem cabeçalhos
        with pd.ExcelWriter(caminho_apenas_valores, engine='openpyxl') as writer:
            # Salvar sem cabeçalhos de coluna e sem índice
            df_apenas_valores.to_excel(writer, sheet_name='Valores', 
                                       header=False, index=False)
        
        print(f"✓ Relatório 1 criado: {caminho_apenas_valores}")
        print(f"  Shape: {df_apenas_valores.shape}")
        
        # ========== RELATÓRIO 2: RELATÓRIO COMPLETO ==========
        print("\n--- CRIANDO RELATÓRIO 2: RELATÓRIO COMPLETO ---")
        
        # Reset index para ter a coluna "Procedimento" no relatório completo
        df_completo = df_consolidado.reset_index()
        df_completo = df_completo.rename(columns={'index': 'Procedimento'})
        
        # Preencher NaN com 0 (para garantir)
        df_completo = df_completo.fillna(0)
        
        # Salvar arquivo consolidado completo
        caminho_completo = "Prestador/prontobaby/resultado/relatorio_prontobaby_COMPLETO.xlsx"
        df_completo.to_excel(caminho_completo, index=False)
        
        print(f"✓ Relatório 2 criado: {caminho_completo}")
        print(f"  Shape do DataFrame: {df_completo.shape}")
        
        # Mostrar preview dos dados completos
        print("\nPreview do relatório completo:")
        print(df_completo.head())
        
        # Mostrar totais por município
        print("\nTotais por município:")
        for municipio in municipios:
            total = totais_municipios[municipio]
            print(f"  {municipio}: {total}")
        
        # Mostrar formato dos dois relatórios
        print("\n--- RESUMO DOS RELATÓRIOS ---")
        print(f"Relatório 1 (apenas valores): {df_apenas_valores.shape[0]} linhas x {df_apenas_valores.shape[1]} colunas")
        print(f"Relatório 2 (completo): {df_completo.shape[0]} linhas x {df_completo.shape[1]} colunas")
        
        # Exemplo de como ficam os dados
        print("\nExemplo dos dados (primeiras 3 linhas):")
        print("Relatório 1 (apenas valores):")
        print(df_apenas_valores.iloc[:3, :3].to_string(header=False, index=False))
        print("\nRelatório 2 (completo):")
        print(df_completo.iloc[:3, :3].to_string(index=False))
    
    # 2. Criar DataFrame consolidado para procedimentos não listados (apenas completo)
    if nao_listados_consolidados:
        # Converter dicionário para DataFrame
        df_nao_listados = pd.DataFrame.from_dict(nao_listados_consolidados, orient='index')
        
        # Garantir que todas as colunas (municípios) existam
        for municipio in municipios:
            if municipio not in df_nao_listados.columns:
                df_nao_listados[municipio] = 0
        
        # Ordenar por soma total (decrescente)
        df_nao_listados['TOTAL'] = df_nao_listados[municipios].sum(axis=1)
        df_nao_listados = df_nao_listados.sort_values('TOTAL', ascending=False)
        
        # Reset index para ter a coluna "Procedimento"
        df_nao_listados = df_nao_listados.reset_index()
        df_nao_listados = df_nao_listados.rename(columns={'index': 'Procedimento'})
        
        # Preencher NaN com 0
        df_nao_listados = df_nao_listados.fillna(0)
        
        # Salvar arquivo consolidado de não listados (apenas completo)
        caminho_nao_listados = "Prestador/prontobaby/resultado/CONSOLIDADO_NAO_LISTADOS.xlsx"
        df_nao_listados.to_excel(caminho_nao_listados, index=False)
        print(f"\n✓ Arquivo consolidado de não listados criado: {caminho_nao_listados}")
        print(f"  Total de procedimentos não listados: {len(df_nao_listados)}")
    else:
        print("\n✗ Nenhum procedimento não listado encontrado")
    
    print(f"\n=== PROCESSAMENTO CONCLUÍDO ===")
    print(f"Total de municípios processados: {len(municipios)}")

# Executar a função
if __name__ == "__main__":
    analisar_prontobaby()