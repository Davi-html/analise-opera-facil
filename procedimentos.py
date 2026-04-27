import pandas as pd

arquivo_excel = "db.xlsx"


def carregar(tabela, nome_coluna):
    if nome_coluna not in tabela.columns:
        return []

    return [
        x for x in tabela[nome_coluna].to_list()
        if pd.notna(x) and x not in [None, '']
    ]


# ORDEM DA SUA PLANILHA

procedimentos = {
    # PEDIÁTRICO - PACOTES
    "pacote_otorrino_pediatrico": "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OTORRINO",
    "pacote_geral_pediatrico": "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO CIRURGIA GERAL",
    "pacote_oftalmo_pediatrico": "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OFTALMOLOGISTA",
    "pacote_risco_cirurgico": "PACOTE RISCO CIRURGICO PEDIÁTRICO",

    # PEDIÁTRICO - PROCEDIMENTOS
    "adenoidectomia": "ADENOIDECTOMIA PEDIÁTRICO",
    "amigdalectomia": "AMIGDALECTOMIA- PEDIATRICO",
    "amigdalectomia_com_adenoide": "AMIGDALECTOMIA COM ADENOIDECTOMIA - PEDIATRICO",
    "tratamento_septo_nasal": "TRATAMENTO CIRÚRGICO DE PERFURAÇÃO DO SEPTO NASAL - PEDIATRICO",
    "correcao_estrabismo": "CORREÇÃO CIRÚRGICA DE ESTRABISMO (ACIMA DE 2 MUSCULOS) - PEDIATRICO",
    "hernia_inguinal_bilateral_ped": "HERNIOPLASTIA INGUINAL (BILATERAL) - PEDIATRICO",
    "hernia_umbilical_ped": "HERNIOPLASTIA UMBILICAL - PEDIATRICO",
    "orquidopexia": "ORQUIDOPEXIA BILATERAL - PEDIATRICO",
    "hidrocele_ped": "TRATAMENTO CIRÚRGICO DE HIDROCELE - PEDIATRICO",
    "hipospadia": "CORRECAO DE HIPOSPADIA (1º TEMPO) - PEDIATRICO",
    "plastica_penis_ped": "PLASTICA TOTAL DO PENIS - PEDIATRICO",
    "postectomia_ped": "POSTECTOMIA - PEDIATRICO",

    # ADULTO - PACOTES
    "pacote_geral_adulto": "PACOTE PRÉ-OPERATÓRIO ADULTO CIRURGIA GERAL",
    "pacote_risco_cirurgico_adulto": "PACOTE RISCO CIRURGICO - ADULTO",
    "pacote_vascular_adulto": "PACOTE PRÉ-OPERATÓRIO ADULTO - VASCULAR",
    "pacote_urologia_adulto": "PACOTE PRÉ-OPERATÓRIO ADULTO - UROLOGIA",
    "pacote_ginecologia_adulto": "PACOTE PRÉ-OPERATÓRIO ADULTO - GINECOLOGIA",
    "pacote_procto_adulto": "PACOTE PRÉ-OPERATÓRIO ADULTO - PROCTOLOGISTA",

    # ADULTO - CIRURGIAS
    "cirurgia_varizes": "CIRURGIA DE VARIZES UNILATERAL (COM SAFENA) - ADULTO",
    "cirurgia_hemorroida": "CIRURGIA DE HEMORROIDECTOMIA - ADULTO",
    "cirurgia_colecistectomia": "CIRURGIA DE COLECISTECTOMIA - ADULTO",
    "cirurgia_colecistectomia_video": "CIRURGIA DE COLECISTECTOMIA VIDEOLAPAROSCÓPICA - ADULTO",
    "cirurgia_hernia_epigastrica": "CIRURGIA DE HERNIOPLASTIA EPIGASTRICA - ADULTO",
    "cirurgia_hernia_incisional": "CIRURGIA DE HERNIOPLASTIA INCISIONAL - ADULTO",
    "cirurgia_hernia_inguinal": "CIRURGIA DE HERNIOPLASTIA INGUINAL (UNILATERAL) - ADULTO",
    "cirurgia_hernia_recidivante": "CIRURGIA DE HERNIOPLASTIA RECIDIVANTE - ADULTO",
    "cirurgia_hernia_umbilical": "CIRURGIA DE HERNIOPLASTIA UMBILICAL - ADULTO",
    "cirurgia_hernia_inguinal_video": "CIRURGIA DE HERNIORRAFIA INGUINAL VIDEOLAPAROSCÓPICA - ADULTO",
    "cirurgia_hernia_umbilical_video": "CIRURGIA DE HERNIORRAFIA UMBILICAL VIDEOLAPAROSCÓPICA - ADULTO",
    "cirurgia_laparotomia": "CIRURGIA DE LAPAROTOMIA EXPLORADORA - ADULTO",
    "cirurgia_duplo_j": "CIRURGIA DE INSTALAÇÃO ENDOSCÓPICA DE CATETER DUPLO J - ADULTO",
    "cirurgia_litotripsia": "CIRURGIA DE LITOTRIPSIA - ADULTO",
    "cirurgia_uretrop": "CIRURGIA DE URETROPLASTIA (RESSECÇÃO DE CORDA) - ADULTO",
    "cirurgia_resseccao_prostata": "CIRURGIA DE RESSECÇÃO ENDOSCÓPICA DE PRÓSTATA - ADULTO",
    "cirurgia_espermatocele": "CIRURGIA DE ESPERMATOCELECTOMIA - ADULTO",
    "cirurgia_hidrocele": "CIRURGIA DE HIDROCELE - ADULTO",
    "cirurgia_varicocele": "CIRURGIA DE VARICOCELE (VARICOCELECTOMIA) - ADULTO",
    "cirurgia_vasectomia": "CIRURGIA DE VASECTOMIA - ADULTO",
    "cirurgia_plastica_penis": "CIRURGIA DE PLASTICA TOTAL DO PENIS - ADULTO",
    "cirurgia_postectomia": "CIRURGIA DE POSTECTOMIA (FIMOSE) - ADULTO",
    "cirurgia_curetagem": "CIRURGIA DE CURETAGEM SEMIOTICA - ADULTO",
    "cirurgia_histerectomia": "CIRURGIA DE HISTERECTOMIA TOTAL - ADULTO",
    "cirurgia_histerectomia_video": "CIRURGIA DE HISTERECTOMIA VIDEOLAPAROSCOPICA - ADULTO",
    "cirurgia_laqueadura": "CIRURGIA DE LAQUEADURA TUBARIA - ADULTO",
    "cirurgia_miomectomia": "CIRURGIA DE MIOMECTOMIA - ADULTO",
    "cirurgia_miomectomia_video": "CIRURGIA DE MIOMECTOMIA VIDEOLAPAROSCOPICA - ADULTO",
    "cirurgia_ooforectomia": "CIRURGIA DE OOFORECTOMIA / OOFOROPLASTIA - ADULTO",
    "cirurgia_salpingectomia": "CIRURGIA DE SALPINGECTOMIA UNI / BILATERAL - ADULTO",
    "cirurgia_colpoperineoplastia": "CIRURGIA DE COLPOPERINEOPLASTIA ANTERIOR E POSTERIOR - ADULTO",
    "cirurgia_cisto_vaginal": "CIRURGIA DE EXERESE DE CISTO VAGINAL - ADULTO",
    "cirurgia_bartholin": "CIRURGIA DE EXERESE DE GLÂNDULA DE BARTHOLIN / SKENE - ADULTO",
    "cirurgia_prostatectomia": "CIRURGIA DE PROSTATECTOMIA - ADULTO",
}