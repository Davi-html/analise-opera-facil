import pandas as pd
import os

def analise_financeiro(competencia, ano, prestadores):

    for prestador in prestadores:

        tabela = pd.read_excel(
            f'relatorios_simplificados/separar{prestador}_SIMPLIFICADO.xlsx',
            sheet_name='Dados Detalhados'
        )

        # 🔥 PADRONIZAÇÃO
        tabela['Procedimento'] = tabela['Procedimento'].astype(str).str.strip().str.upper()
        tabela['municipio'] = tabela['Municipio'].astype(str).str.replace('RJ - ', '', regex=False)

        # 🔥 FILTRAR SÓ ADULTO
        tabela = tabela[tabela['Procedimento'].str.contains('ADULTO', na=False)]

        municipios = sorted(tabela['municipio'].unique())

        valor_unitario_cirurgia = {
            "CIRURGIA DE VARIZES UNILATERAL (COM SAFENA) - ADULTO": 8132.00,
            "CIRURGIA DE HEMORROIDECTOMIA - ADULTO": 5850.00,
            "CIRURGIA DE COLECISTECTOMIA - ADULTO": 10362.30,
            "CIRURGIA DE COLECISTECTOMIA VIDEOLAPAROSCÓPICA - ADULTO": 13100.00,
            "CIRURGIA DE HERNIOPLASTIA EPIGASTRICA - ADULTO": 5850.00,
            "CIRURGIA DE HERNIOPLASTIA INCISIONAL - ADULTO": 6175.00,
            "CIRURGIA DE HERNIOPLASTIA INGUINAL (UNILATERAL) - ADULTO": 5850.00,
            "CIRURGIA DE HERNIOPLASTIA RECIDIVANTE - ADULTO": 7312.50,
            "CIRURGIA DE HERNIOPLASTIA UMBILICAL - ADULTO": 6175.00,
            "CIRURGIA DE HERNIORRAFIA INGUINAL VIDEOLAPAROSCÓPICA - ADULTO": 8626.00,
            "CIRURGIA DE HERNIORRAFIA UMBILICAL VIDEOLAPAROSCÓPICA - ADULTO": 8626.00,
            "CIRURGIA DE LAPAROTOMIA EXPLORADORA - ADULTO": 10838.23,
            "CIRURGIA DE INSTALAÇÃO ENDOSCÓPICA DE CATETER DUPLO J - ADULTO": 5173.36,
            "CIRURGIA DE LITOTRIPSIA - ADULTO": 13750.00,
            "CIRURGIA DE URETROPLASTIA (RESSECÇÃO DE CORDA) - ADULTO": 13500.00,
            "CIRURGIA DE RESSECÇÃO ENDOSCÓPICA DE PRÓSTATA - ADULTO": 14500.00,
            "CIRURGIA DE ESPERMATOCELECTOMIA - ADULTO": 4850.00,
            "CIRURGIA DE HIDROCELE - ADULTO": 4850.00,
            "CIRURGIA DE VARICOCELE (VARICOCELECTOMIA) - ADULTO": 5082.80,
            "CIRURGIA DE VASECTOMIA - ADULTO": 5850.00,
            "CIRURGIA DE PLASTICA TOTAL DO PENIS - ADULTO": 12500.00,
            "CIRURGIA DE POSTECTOMIA (FIMOSE) - ADULTO": 5850.00,
            "CIRURGIA DE CURETAGEM SEMIOTICA - ADULTO": 5850.00,
            "CIRURGIA DE HISTERECTOMIA TOTAL - ADULTO": 10886.20,
            "CIRURGIA DE HISTERECTOMIA VIDEOLAPAROSCOPICA - ADULTO": 16340.25,
            "CIRURGIA DE LAQUEADURA TUBARIA - ADULTO": 6500.00,
            "CIRURGIA DE MIOMECTOMIA - ADULTO": 8500.00,
            "CIRURGIA DE MIOMECTOMIA VIDEOLAPAROSCOPICA - ADULTO": 16125.00,
            "CIRURGIA DE OOFORECTOMIA / OOFOROPLASTIA - ADULTO": 8500.00,
            "CIRURGIA DE SALPINGECTOMIA UNI / BILATERAL - ADULTO": 8500.00,
            "CIRURGIA DE COLPOPERINEOPLASTIA ANTERIOR E POSTERIOR - ADULTO": 16500.00,
            "CIRURGIA DE EXERESE DE CISTO VAGINAL - ADULTO": 4000.00,
            "CIRURGIA DE EXERESE DE GLÂNDULA DE BARTHOLIN / SKENE - ADULTO": 3500.00,
            "CIRURGIA DE PROSTATECTOMIA - ADULTO": 8385.29
        }

        valor_unitario_pacote = {
            "PACOTE PRÉ-OPERATÓRIO ADULTO - GINECOLOGIA": 450,
            "PACOTE PRÉ-OPERATÓRIO ADULTO - PROCTOLOGISTA": 450,
            "PACOTE PRÉ-OPERATÓRIO ADULTO - UROLOGIA": 450,
            "PACOTE PRÉ-OPERATÓRIO ADULTO - VASCULAR": 450,
            "PACOTE PRÉ-OPERATÓRIO ADULTO CIRURGIA GERAL": 450,
            "PACOTE RISCO CIRURGICO - ADULTO": 400
        }

        dados = []

        for municipio in municipios:
            dados_municipio = tabela[tabela['municipio'] == municipio]

            # CIRURGIAS
            for nome, valor in valor_unitario_cirurgia.items():
                qtd = dados_municipio.loc[
                    dados_municipio['Procedimento'] == nome, 'Quantidade'
                ].sum()

                dados.append({
                    'Prestador': prestador,
                    'Tipo': 'CIRURGIA',
                    'Procedimento': nome,
                    'Quantidade': qtd,
                    'Valor Unitario': valor,
                    'Total': qtd * valor,
                    'Municipio': municipio,
                    'Ano': ano,
                    'Competencia': competencia
                })

            # PACOTES
            for nome, valor in valor_unitario_pacote.items():
                qtd = dados_municipio.loc[
                    dados_municipio['Procedimento'] == nome, 'Quantidade'
                ].sum()

                dados.append({
                    'Prestador': prestador,
                    'Tipo': 'PACOTE',
                    'Procedimento': nome,
                    'Quantidade': qtd,
                    'Valor Unitario': valor,
                    'Total': qtd * valor,
                    'Municipio': municipio,
                    'Ano': ano,
                    'Competencia': competencia
                })

        df = pd.DataFrame(dados)

        os.makedirs(f'analise_financeiro/{prestador}', exist_ok=True)
        df.to_excel(f'analise_financeiro/{prestador}/adulto.xlsx', index=False)

        print(f"✅ {prestador} - Cirurgias + Pacotes OK")
