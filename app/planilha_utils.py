import math
from html import escape

import pandas as pd


COLUNAS_OBRIGATORIAS = [
    "titulo",
    "descricao",
    "categoria_id",
    "localizacao_id",
    "tecnico_id",
    "requerente_id",
]


def normalizar_colunas(df):
    df.columns = [str(col).strip().lower() for col in df.columns]
    return df


def is_empty(value):
    if pd.isna(value):
        return True
    texto = str(value).strip()
    return texto == "" or texto.lower() == "nan"


def int_or_none(value):
    if pd.isna(value):
        return None
    try:
        if isinstance(value, float) and math.isnan(value):
            return None
        return int(float(value))
    except Exception:
        return None


def preparar_texto_glpi(value, usar_html=False):
    if pd.isna(value):
        return ""

    texto = str(value).replace("\r\n", "\n").replace("\r", "\n")
    if not usar_html:
        return texto

    # Converte texto livre em HTML basico para preservar quebra de linha e listas.
    linhas = texto.split("\n")
    partes = []
    em_lista = False

    for linha in linhas:
        limpa = linha.strip()
        if limpa.startswith("- ") or limpa.startswith("* "):
            if not em_lista:
                partes.append("<ul>")
                em_lista = True
            item = escape(limpa[2:].strip())
            partes.append(f"<li>{item}</li>")
            continue

        if em_lista:
            partes.append("</ul>")
            em_lista = False

        if limpa == "":
            partes.append("<br>")
            continue

        # Preserva espacos duplos e tabs em HTML.
        linha_html = escape(linha).replace("  ", "&nbsp;&nbsp;").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;")
        partes.append(linha_html + "<br>")

    if em_lista:
        partes.append("</ul>")

    return "".join(partes).rstrip("<br>")


def validar_dataframe(df):
    colunas_faltantes = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]

    total = len(df)
    linhas_invalidas = []
    validas = 0

    if not colunas_faltantes:
        for index, row in df.iterrows():
            linha_excel = index + 2
            titulo = row["titulo"]
            categoria_id = row["categoria_id"]
            localizacao_id = row["localizacao_id"]
            tecnico_id = row["tecnico_id"]
            requerente_id = row["requerente_id"]

            if is_empty(titulo):
                linhas_invalidas.append((linha_excel, "titulo vazio"))
                continue

            campos_numericos = {
                "categoria_id": categoria_id,
                "localizacao_id": localizacao_id,
                "tecnico_id": tecnico_id,
                "requerente_id": requerente_id,
            }
            campo_invalido = None
            for nome, valor in campos_numericos.items():
                if int_or_none(valor) is None:
                    campo_invalido = nome
                    break

            if campo_invalido:
                linhas_invalidas.append((linha_excel, f"{campo_invalido} invalido"))
                continue

            validas += 1

    invalidas = total - validas if not colunas_faltantes else total
    return {
        "colunas_faltantes": colunas_faltantes,
        "linhas_invalidas": linhas_invalidas,
        "total": total,
        "validas": validas,
        "invalidas": invalidas,
    }
