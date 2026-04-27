import pandas as pd
from collections import Counter

# Caminhos dos arquivos - ATENÇÃO: tem um erro de digitação "ralatorio" em vez de "relatorio"
faltaNeotin = "./relatorios_simplificados/neotin-falta.xlsx"  # CORRIGIDO
relatorioNeotin = "./relatorios_simplificados/separarNeotin_SIMPLIFICADO.xlsx"     # CORRIGIDO

MUNICIPIOS = ["RJ - Belford Roxo", "RJ - Duque de Caxias", "RJ - Itaguaí", "RJ - Japeri", 
                  "RJ - Magé", "RJ - Mesquita", "RJ - Nilópolis", "RJ - Nova Iguaçu", 
                  "RJ - Paracambi", "RJ - Queimados", "RJ - Seropédica", "RJ - São João de Meriti"]


def ler_todos_pacientes_relatorio(caminho):
    """Lê todos os pacientes do relatório, considerando todos os municípios"""
    try:
        df = pd.read_excel(caminho)
        print(f"✅ Arquivo {caminho} carregado com sucesso!")
        print(f"   Linhas: {df.shape[0]}, Colunas: {df.shape[1]}")
        
        # Mostrar colunas disponíveis
        print("\n📋 Colunas disponíveis no relatório:")
        colunas_pacientes = [col for col in df.columns if 'Paciente' in col]
        for col in colunas_pacientes:
            print(f"   - {col}")
        
    except FileNotFoundError:
        print(f"❌ Arquivo {caminho} não encontrado!")
        print("   Verifique se o caminho está correto:")
        print(f"   Caminho atual: {caminho}")
        return []
    except Exception as e:
        print(f"❌ Erro ao ler {caminho}: {e}")
        return []

    # Coletar todos os pacientes de todas as colunas de municípios
    todos_pacientes = []
    detalhes_pacientes = []  # Para armazenar informações detalhadas
    
    for municipio in MUNICIPIOS:
        coluna_paciente = f"Paciente {municipio}"
        
        if coluna_paciente in df.columns:
            # Verificar se há coluna de quantidade também
            coluna_quantidade = f"Quantidade {municipio}"
            tem_quantidade = coluna_quantidade in df.columns
            
            for idx, valor in enumerate(df[coluna_paciente]):
                if pd.notna(valor) and str(valor).strip() not in ('', 'nan'):
                    paciente = str(valor).strip()
                    todos_pacientes.append(paciente)
                    
                    # Armazenar detalhes
                    detalhe = {
                        'paciente': paciente,
                        'municipio': municipio,
                        'linha': idx + 2,  # +2 porque Excel começa em 1 e o header é linha 1
                        'coluna': coluna_paciente
                    }
                    
                    # Adicionar quantidade se existir
                    if tem_quantidade:
                        detalhe['quantidade'] = df[coluna_quantidade].iloc[idx] if idx < len(df) else None
                    
                    detalhes_pacientes.append(detalhe)
    
    return todos_pacientes, detalhes_pacientes


def analisar_duplicatas(pacientes, detalhes):
    """Analisa pacientes duplicados no relatório"""
    
    print("\n" + "="*70)
    print("ANÁLISE DE PACIENTES DUPLICADOS NO RELATÓRIO")
    print("="*70)
    
    # Contar frequência de cada paciente
    contador = Counter(pacientes)
    
    # Separar únicos e duplicados
    pacientes_unicos = [p for p, c in contador.items() if c == 1]
    pacientes_duplicados = [p for p, c in contador.items() if c > 1]
    
    print(f"\n📊 ESTATÍSTICAS GERAIS:")
    print(f"   • Total de pacientes no relatório: {len(pacientes)}")
    print(f"   • Pacientes distintos: {len(contador)}")
    print(f"   • Pacientes únicos (aparecem 1 vez): {len(pacientes_unicos)}")
    print(f"   • Pacientes duplicados (aparecem 2+ vezes): {len(pacientes_duplicados)}")
    
    # Mostrar contagem por frequência
    print(f"\n📈 DISTRIBUIÇÃO DE FREQUÊNCIA:")
    for freq in sorted(set(contador.values())):
        quantidade = len([p for p, c in contador.items() if c == freq])
        print(f"   • Aparecem {freq} vez(es): {quantidade} paciente(s)")
    
    # Análise detalhada dos duplicados
    if pacientes_duplicados:
        print(f"\n🔍 PACIENTES DUPLICADOS (aparecem 2 ou mais vezes):")
        print("-" * 70)
        
        for paciente in sorted(pacientes_duplicados):
            frequencia = contador[paciente]
            
            # Encontrar todos os registros deste paciente
            registros = [d for d in detalhes if d['paciente'] == paciente]
            
            print(f"\n📌 {paciente}")
            print(f"   Aparece {frequencia} vez(es) no relatório:")
            
            for i, registro in enumerate(registros, 1):
                municipio = registro['municipio']
                linha = registro['linha']
                coluna = registro['coluna']
                quantidade = registro.get('quantidade', 'N/A')
                
                print(f"   {i}. Município: {municipio}")
                print(f"      Linha Excel: {linha}")
                print(f"      Coluna: {coluna}")
                if quantidade != 'N/A':
                    print(f"      Quantidade: {quantidade}")
    
    # Análise de pacientes que aparecem em múltiplos municípios
    print(f"\n🌍 PACIENTES EM MÚLTIPLOS MUNICÍPIOS:")
    print("-" * 70)
    
    # Agrupar pacientes por município
    pacientes_por_municipio = {}
    for detalhe in detalhes:
        paciente = detalhe['paciente']
        municipio = detalhe['municipio']
        
        if paciente not in pacientes_por_municipio:
            pacientes_por_municipio[paciente] = set()
        pacientes_por_municipio[paciente].add(municipio)
    
    # Encontrar pacientes em múltiplos municípios
    pacientes_mult_municipios = {p: muns for p, muns in pacientes_por_municipio.items() if len(muns) > 1}
    
    if pacientes_mult_municipios:
        print(f"   {len(pacientes_mult_municipios)} paciente(s) aparecem em mais de um município:")
        
        for paciente, municipios in sorted(pacientes_mult_municipios.items(), key=lambda x: len(x[1]), reverse=True):
            print(f"\n   📌 {paciente}")
            print(f"      Municípios: {', '.join(sorted(municipios))}")
            print(f"      Total de municípios: {len(municipios)}")
            
            # Mostrar detalhes de cada ocorrência
            registros = [d for d in detalhes if d['paciente'] == paciente]
            for registro in registros:
                print(f"      • {registro['municipio']} (Linha {registro['linha']})")
    else:
        print("   Nenhum paciente aparece em múltiplos municípios.")
    
    return contador, pacientes_unicos, pacientes_duplicados, pacientes_mult_municipios


def exportar_resultados(contador, detalhes, pacientes_mult_municipios):
    """Exporta os resultados para Excel"""
    
    # Preparar dados para exportação
    dados_exportacao = []
    
    for detalhe in detalhes:
        paciente = detalhe['paciente']
        frequencia = contador[paciente]
        
        dados_exportacao.append({
            'Paciente': paciente,
            'Município': detalhe['municipio'],
            'Frequência no Relatório': frequencia,
            'Linha Excel': detalhe['linha'],
            'Coluna': detalhe['coluna'],
            'Quantidade': detalhe.get('quantidade', 'N/A'),
            'É Duplicado?': 'SIM' if frequencia > 1 else 'NÃO',
            'Aparece em Múltiplos Municípios?': 'SIM' if paciente in pacientes_mult_municipios else 'NÃO',
            'Municípios (se múltiplos)': ', '.join(pacientes_mult_municipios.get(paciente, [])) if paciente in pacientes_mult_municipios else ''
        })
    
    # Criar DataFrame e exportar
    df_export = pd.DataFrame(dados_exportacao)
    
    # Ordenar por frequência (mais duplicados primeiro)
    df_export = df_export.sort_values(['Frequência no Relatório', 'Paciente', 'Município'], 
                                      ascending=[False, True, True])
    
    # Exportar para Excel
    nome_arquivo = "analise_duplicatas_relatorio.xlsx"
    
    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        # Aba com todos os dados
        df_export.to_excel(writer, sheet_name='Todos_Registros', index=False)
        
        # Aba apenas com duplicados
        duplicados = df_export[df_export['Frequência no Relatório'] > 1]
        duplicados.to_excel(writer, sheet_name='Apenas_Duplicados', index=False)
        
        # Aba com resumo estatístico
        resumo_data = {
            'Métrica': [
                'Total de registros de pacientes',
                'Pacientes distintos',
                'Pacientes únicos (1 ocorrência)',
                'Pacientes duplicados (2+ ocorrências)',
                'Pacientes em múltiplos municípios'
            ],
            'Valor': [
                len(dados_exportacao),
                len(contador),
                len([p for p, c in contador.items() if c == 1]),
                len([p for p, c in contador.items() if c > 1]),
                len(pacientes_mult_municipios)
            ]
        }
        df_resumo = pd.DataFrame(resumo_data)
        df_resumo.to_excel(writer, sheet_name='Resumo_Estatistico', index=False)
        
        # Aba com top 20 mais duplicados
        top_duplicados = df_export[['Paciente', 'Frequência no Relatório']].drop_duplicates()
        top_duplicados = top_duplicados.sort_values('Frequência no Relatório', ascending=False).head(20)
        top_duplicados.to_excel(writer, sheet_name='Top_20_Duplicados', index=False)
    
    print(f"\n💾 Resultados exportados para: {nome_arquivo}")
    print("   Abas do arquivo:")
    print("   1. Todos_Registros - Lista completa de todos os pacientes")
    print("   2. Apenas_Duplicados - Somente pacientes que se repetem")
    print("   3. Resumo_Estatistico - Estatísticas gerais")
    print("   4. Top_20_Duplicados - 20 pacientes mais duplicados")


def main():
    """Função principal"""
    
    print("=" * 70)
    print("ANÁLISE DE DUPLICATAS NO RELATÓRIO DE PACIENTES")
    print("=" * 70)
    
    # Ler dados do relatório
    pacientes, detalhes = ler_todos_pacientes_relatorio(relatorioNeotin)
    
    if not pacientes:
        print("\n❌ Não foi possível carregar os dados. Verifique o arquivo.")
        return
    
    # Analisar duplicatas
    contador, unicos, duplicados, mult_municipios = analisar_duplicatas(pacientes, detalhes)
    
    # Exportar resultados
    exportar_resultados(contador, detalhes, mult_municipios)
    
    print("\n" + "=" * 70)
    print("ANÁLISE CONCLUÍDA!")
    print("=" * 70)


if __name__ == '__main__':
    main()