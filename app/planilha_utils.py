import math
import unicodedata
from html import escape

import pandas as pd


COLUNAS_OBRIGATORIAS = [
    "titulo",
    "descricao",
    "categoria_id",
    "localizacao_id",
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


def _normalizar_texto_tipo(value):
    texto = str(value).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return texto


def tipo_ticket_or_none(value):
    if is_empty(value):
        return None

    numero = int_or_none(value)
    if numero in (1, 2):
        return numero

    texto = _normalizar_texto_tipo(value)
    mapa = {
        "incidente": 1,
        "incident": 1,
        "requisicao": 2,
        "requisicao de servico": 2,
        "request": 2,
        "solicitacao": 2,
    }
    return mapa.get(texto)


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
                partes.append("<ul>\n")
                em_lista = True
            item = escape(limpa[2:].strip())
            partes.append(f"<li>{item}</li>\n")
            continue

        if em_lista:
            partes.append("</ul>\n")
            em_lista = False

        if limpa == "":
            partes.append("<p>&nbsp;</p>\n")
            continue

        # Preserva espacos duplos e tabs em HTML.
        linha_html = escape(linha).replace("  ", "&nbsp;&nbsp;").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;")
        partes.append(f"<p>{linha_html}</p>\n")

    if em_lista:
        partes.append("</ul>\n")

    return "".join(partes).rstrip()


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
            tecnico_id = row.get("tecnico_id")
            requerente_id = row.get("requerente_id")
            tipo = row.get("tipo")

            if is_empty(titulo):
                linhas_invalidas.append((linha_excel, "titulo vazio"))
                continue

            campos_numericos_obrigatorios = {
                "categoria_id": categoria_id,
                "localizacao_id": localizacao_id,
            }
            campo_invalido = None
            for nome, valor in campos_numericos_obrigatorios.items():
                if int_or_none(valor) is None:
                    campo_invalido = nome
                    break

            if campo_invalido:
                linhas_invalidas.append((linha_excel, f"{campo_invalido} invalido"))
                continue

            campos_numericos_opcionais = {
                "tecnico_id": tecnico_id,
                "requerente_id": requerente_id,
            }
            for nome, valor in campos_numericos_opcionais.items():
                if is_empty(valor):
                    continue
                if int_or_none(valor) is None:
                    campo_invalido = nome
                    break

            if campo_invalido:
                linhas_invalidas.append((linha_excel, f"{campo_invalido} invalido"))
                continue

            if not is_empty(tipo) and tipo_ticket_or_none(tipo) is None:
                linhas_invalidas.append((linha_excel, "tipo invalido (use incidente/requisicao ou 1/2)"))
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
