import pandas as pd

def analise_financeiro(competencia, ano):
    for prestador in ['Neomater', 'Neotin', 'Pediatrico']:
    # Carregar o arquivo Excel original
        tabela = pd.read_excel('relatorios_simplificados/separar{}_SIMPLIFICADO.xlsx'.format(prestador),
                            sheet_name='Dados Detalhados')

        # Dicionários de valores unitários
        valor_unitario_cirurgia = {
            "ADENOIDECTOMIA PEDIÁTRICO": 5330,
            "AMIGDALECTOMIA- PEDIATRICO": 6713.01,
            "AMIGDALECTOMIA COM ADENOIDECTOMIA - PEDIATRICO": 7698.35,
            "TRATAMENTO CIRÚRGICO DE PERFURAÇÃO DO SEPTO NASAL - PEDIATRICO": 6500,
            "CORREÇÃO CIRÚRGICO DE ESTRABISMO (ACIMA DE 2 MUSCULOS) - PEDIATRICO": 5255.28,
            "HERNIOPLASTIA INGUINAL (BILATERAL) - PEDIATRICO": 5850,
            "HERNIOPLASTIA UMBILICAL - PEDIATRICO": 5237.06,
            "ORQUIDOPEXIA BILATERAL - PEDIATRICO": 7157.78,
            "TRATAMENTO CIRÚRGICO DE HIDROCELE - PEDIATRICO": 3782.7,
            "CORRECAO DE HIPOSPADIA (1º TEMPO) - PEDIATRICO": 6608.86,
            "PLASTICA TOTAL DO PENIS - PEDIATRICO": 6500,
            "POSTECTOMIA - PEDIATRICO": 4850

        }

        valor_unitario_pacote = {
            "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OTORRINO": 300,
            "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO CIRURGIA GERAL": 300,
            "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OFTALMOLOGISTA": 300  
        }


        tabela['municipio'] = tabela['Municipio'].astype(str).str.replace('RJ - ', '', regex=False)

        municipios = sorted(tabela['municipio'].unique())

        # Criar lista para os dados
        dados = []

        # Para cada município
        for municipio in municipios:
            # Filtrar dados do município
            dados_municipio = tabela[tabela['municipio'] == municipio]
            
            # Para CADA cirurgia - uma linha por cirurgia
            for cirurgia_nome, cirurgia_valor in valor_unitario_cirurgia.items():
                # Verificar se esta cirurgia ocorreu
                cirurgia_ocorreu = dados_municipio[dados_municipio['Procedimento'] == cirurgia_nome]
                

                quantidade_cirurgia = 0
                valor_total_cirurgia = 0

                if not cirurgia_ocorreu.empty:
                    quantidade_cirurgia = cirurgia_ocorreu['Quantidade'].sum()
                    valor_total_cirurgia = quantidade_cirurgia * cirurgia_valor
                
                # Adicionar registro da cirurgia
                dados.append({
                    'Prestador': prestador,
                    'Cirurgias': cirurgia_nome,
                    'Valor Unitário': cirurgia_valor,
                    'Quantidade': quantidade_cirurgia,
                    'MUNICIPIO': municipio,
                    'Ano': ano,
                    'Competencia': competencia,
                    'total gasto cirurgia': valor_total_cirurgia,
                })
            
            # Para CADA consulta - uma linha por consulta
            for consulta_nome, consulta_valor in valor_unitario_pacote.items():
                # Verificar se esta consulta ocorreu
                consulta_ocorreu = dados_municipio[dados_municipio['Procedimento'] == consulta_nome]
                
                quantidade_consulta = 0
                valor_total_consulta = 0

                if not consulta_ocorreu.empty:
                    quantidade_consulta = consulta_ocorreu['Quantidade'].sum()
                    valor_total_consulta = quantidade_consulta * consulta_valor
                
                consulta_nome = consulta_nome.replace("PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO ", "CONSULTA PEDIATRICA ")
                
                # Adicionar registro da consulta
                dados.append({
                    'Prestador': prestador,
                    'Consultas': consulta_nome,
                    'Valor Unitário Consulta': consulta_valor,
                    'quantidade_consulta': quantidade_consulta,
                    'MUNICIPIO': municipio,
                    'Ano': ano,
                    'Competencia': competencia,
                    'total gasto consulta': valor_total_consulta,
                })
            
        # Criar DataFrame
        df_resultado = pd.DataFrame(dados)

        # Verificar duplicatas específicas
        df_resultado[['MUNICIPIO', 'Consultas']].drop_duplicates()
        df_resultado[['MUNICIPIO', 'Cirurgias']].drop_duplicates()


        # Salvar
        df_resultado.to_excel('analise_financeiro/{}/apresentação-{}.xlsx'.format(prestador, prestador), index=False)