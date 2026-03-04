import json
from datetime import datetime
from pathlib import Path

import pandas as pd

from .glpi_service import GLPIClient
from .planilha_utils import (
    int_or_none,
    is_empty,
    normalizar_colunas,
    preparar_texto_glpi,
    tipo_ticket_or_none,
    validar_dataframe,
)


class ChamadosBackend:
    def __init__(self, api_url_padrao, config_path):
        self.api_url_padrao = api_url_padrao
        self.config_path = Path(config_path)

        self.df = None
        self.caminho_arquivo = ""
        self.linhas_invalidas = []
        self.colunas_faltantes = []
        self.autenticado = False
        self.api_url = api_url_padrao or ""
        self.app_token = ""
        self.user_token = ""
        self.cliente = None

        self._limpar_referencias_api()

    def _limpar_referencias_api(self):
        self.referencias_api = {
            "usuarios": {},
            "categorias": {},
            "localizacoes": {},
        }
        self.referencias_invalidas = {
            "usuarios": set(),
            "categorias": set(),
            "localizacoes": set(),
        }

    @staticmethod
    def ler_planilha(caminho):
        caminho_lower = caminho.lower()

        if caminho_lower.endswith(".ods"):
            return pd.read_excel(caminho, engine="odf")
        if caminho_lower.endswith(".xls"):
            return pd.read_excel(caminho, engine="xlrd")
        if caminho_lower.endswith((".xlsx", ".xlsm")):
            return pd.read_excel(caminho, engine="openpyxl")

        return pd.read_excel(caminho)

    def carregar_planilha_importacao(self, caminho):
        df = self.ler_planilha(caminho)
        self.df = normalizar_colunas(df)
        self.caminho_arquivo = caminho
        self._limpar_referencias_api()
        return self.df

    def validar_planilha_atual(self):
        if self.df is None:
            return None
        resultado = validar_dataframe(self.df)
        self.colunas_faltantes = resultado["colunas_faltantes"]
        self.linhas_invalidas = resultado["linhas_invalidas"]
        return resultado

    def autenticar(self, url, user_token, app_token):
        cliente = GLPIClient(url, app_token, user_token)
        cliente.autenticar()

        self.api_url = url
        self.user_token = user_token
        self.app_token = app_token
        self.cliente = cliente
        self.autenticado = True

    def carregar_config_local(self):
        if not self.config_path.exists():
            return {}
        try:
            with self.config_path.open("r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def salvar_config_local(self, api_url, user_token, app_token, salvar_tokens):
        payload = {
            "api_url": api_url.strip(),
            "salvar_tokens": bool(salvar_tokens),
            "salvar_user_token": bool(salvar_tokens),
            "user_token": user_token.strip() if salvar_tokens else "",
            "app_token": app_token.strip() if salvar_tokens else "",
        }
        with self.config_path.open("w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=True, indent=2)

    def _nome_por_id(self, valor, mapa):
        item_id = int_or_none(valor)
        if item_id is None:
            return ""
        return mapa.get(item_id, "")

    def construir_preview_df(self, limite=100):
        if self.df is None:
            return None
        preview_df = self.df.head(limite).copy()
        if self.referencias_api["usuarios"] and "tecnico_id" in preview_df.columns:
            preview_df["tecnico_nome"] = preview_df["tecnico_id"].apply(
                lambda x: self._nome_por_id(x, self.referencias_api["usuarios"])
            )
        if self.referencias_api["usuarios"] and "requerente_id" in preview_df.columns:
            preview_df["requerente_nome"] = preview_df["requerente_id"].apply(
                lambda x: self._nome_por_id(x, self.referencias_api["usuarios"])
            )
        if self.referencias_api["categorias"]:
            preview_df["categoria_nome"] = preview_df["categoria_id"].apply(
                lambda x: self._nome_por_id(x, self.referencias_api["categorias"])
            )
        if self.referencias_api["localizacoes"]:
            preview_df["localizacao_nome"] = preview_df["localizacao_id"].apply(
                lambda x: self._nome_por_id(x, self.referencias_api["localizacoes"])
            )
        return preview_df

    def _coletar_ids_referencias(self):
        ids = {
            "usuarios": set(),
            "categorias": set(),
            "localizacoes": set(),
        }
        for _, row in self.df.iterrows():
            tecnico_id = int_or_none(row.get("tecnico_id"))
            requerente_id = int_or_none(row.get("requerente_id"))
            categoria_id = int_or_none(row["categoria_id"])
            localizacao_id = int_or_none(row["localizacao_id"])
            if tecnico_id is not None:
                ids["usuarios"].add(tecnico_id)
            if requerente_id is not None:
                ids["usuarios"].add(requerente_id)
            if categoria_id is not None:
                ids["categorias"].add(categoria_id)
            if localizacao_id is not None:
                ids["localizacoes"].add(localizacao_id)
        return ids

    def _carregar_referencias_api(self, headers, log_cb):
        self._limpar_referencias_api()
        ids = self._coletar_ids_referencias()

        for user_id in sorted(ids["usuarios"]):
            try:
                nome = self.cliente.obter_nome_usuario(headers, user_id)
            except Exception:
                nome = None
            if nome is None:
                self.referencias_invalidas["usuarios"].add(user_id)
                log_cb(f"[AVISO] Usuario {user_id} nao encontrado no GLPI.")
            else:
                self.referencias_api["usuarios"][user_id] = nome

        for categoria_id in sorted(ids["categorias"]):
            try:
                nome = self.cliente.obter_nome_categoria(headers, categoria_id)
            except Exception:
                nome = None
            if nome is None:
                self.referencias_invalidas["categorias"].add(categoria_id)
                log_cb(f"[AVISO] Categoria {categoria_id} nao encontrada no GLPI.")
            else:
                self.referencias_api["categorias"][categoria_id] = nome

        for localizacao_id in sorted(ids["localizacoes"]):
            try:
                nome = self.cliente.obter_nome_localizacao(headers, localizacao_id)
            except Exception:
                nome = None
            if nome is None:
                self.referencias_invalidas["localizacoes"].add(localizacao_id)
                log_cb(f"[AVISO] Localizacao {localizacao_id} nao encontrada no GLPI.")
            else:
                self.referencias_api["localizacoes"][localizacao_id] = nome

    def resumo_referencias_api(self):
        return (
            "Validacao API: "
            f"usuarios {len(self.referencias_api['usuarios'])} ok/{len(self.referencias_invalidas['usuarios'])} invalidos, "
            f"categorias {len(self.referencias_api['categorias'])} ok/{len(self.referencias_invalidas['categorias'])} invalidas, "
            f"localizacoes {len(self.referencias_api['localizacoes'])} ok/{len(self.referencias_invalidas['localizacoes'])} invalidas."
        )

    def consultar_nomes_api(self, log_cb):
        headers = self.cliente.iniciar_sessao()
        try:
            self._carregar_referencias_api(headers, log_cb)
        finally:
            self.cliente.finalizar_sessao(headers)
        return self.resumo_referencias_api()

    def importar_chamados(self, validar_api, usar_html, log_cb, progresso_cb):
        total = len(self.df)
        headers = self.cliente.iniciar_sessao()
        log_cb("[OK] Sessao GLPI iniciada.")

        try:
            if validar_api:
                self._carregar_referencias_api(headers, log_cb)

            sucesso = 0
            erro_api = 0
            ignorados = 0
            detalhes = []

            for index, row in self.df.iterrows():
                linha_excel = index + 2
                titulo = row["titulo"]
                descricao = row["descricao"]

                if is_empty(titulo):
                    ignorados += 1
                    log_cb(f"[AVISO] Linha {linha_excel}: titulo vazio. Ignorada.")
                    detalhes.append(
                        {
                            "linha_excel": linha_excel,
                            "status": "ignorado",
                            "motivo": "titulo vazio",
                        }
                    )
                    progresso_cb(index + 1, total)
                    continue

                categoria_id = int_or_none(row["categoria_id"])
                localizacao_id = int_or_none(row["localizacao_id"])
                tecnico_id = int_or_none(row.get("tecnico_id"))
                requerente_id = int_or_none(row.get("requerente_id"))
                tipo = tipo_ticket_or_none(row.get("tipo"))

                if None in (categoria_id, localizacao_id):
                    ignorados += 1
                    log_cb(f"[AVISO] Linha {linha_excel}: IDs invalidos. Ignorada.")
                    detalhes.append(
                        {
                            "linha_excel": linha_excel,
                            "status": "ignorado",
                            "motivo": "categoria_id/localizacao_id invalido",
                        }
                    )
                    progresso_cb(index + 1, total)
                    continue

                if not is_empty(row.get("tipo")) and tipo is None:
                    ignorados += 1
                    log_cb(f"[AVISO] Linha {linha_excel}: tipo invalido. Use incidente/requisicao ou 1/2. Ignorada.")
                    detalhes.append(
                        {
                            "linha_excel": linha_excel,
                            "status": "ignorado",
                            "motivo": "tipo invalido",
                        }
                    )
                    progresso_cb(index + 1, total)
                    continue

                motivos_invalidos = []
                if validar_api:
                    if tecnico_id is not None and tecnico_id in self.referencias_invalidas["usuarios"]:
                        motivos_invalidos.append("tecnico_id")
                    if requerente_id is not None and requerente_id in self.referencias_invalidas["usuarios"]:
                        motivos_invalidos.append("requerente_id")
                    if categoria_id in self.referencias_invalidas["categorias"]:
                        motivos_invalidos.append("categoria_id")
                    if localizacao_id in self.referencias_invalidas["localizacoes"]:
                        motivos_invalidos.append("localizacao_id")

                if validar_api and motivos_invalidos:
                    ignorados += 1
                    log_cb(
                        f"[AVISO] Linha {linha_excel}: {', '.join(motivos_invalidos)} nao encontrado(s) no GLPI. Ignorada."
                    )
                    detalhes.append(
                        {
                            "linha_excel": linha_excel,
                            "status": "ignorado",
                            "motivo": f"referencias invalidas: {', '.join(motivos_invalidos)}",
                        }
                    )
                    progresso_cb(index + 1, total)
                    continue

                payload_input = {
                    "name": str(titulo).strip(),
                    "content": preparar_texto_glpi(descricao, usar_html=usar_html),
                    "status": 1,
                    "type": tipo or 1,
                    "itilcategories_id": categoria_id,
                    "locations_id": localizacao_id,
                }
                if tecnico_id is not None:
                    payload_input["_users_id_assign"] = tecnico_id
                if requerente_id is not None:
                    payload_input["_users_id_requester"] = requerente_id

                payload = {"input": payload_input}

                try:
                    r = self.cliente.criar_chamado(headers, payload)
                    if r.status_code == 201:
                        sucesso += 1
                        ticket_id = None
                        try:
                            ticket_id = r.json().get("id")
                        except Exception:
                            ticket_id = None
                        detalhes.append(
                            {
                                "linha_excel": linha_excel,
                                "status": "criado",
                                "ticket_id": ticket_id,
                            }
                        )
                    else:
                        erro_api += 1
                        erro_msg = f"status {r.status_code} - {r.text[:180]}"
                        log_cb(f"[ERRO] Linha {linha_excel}: {erro_msg}")
                        detalhes.append(
                            {
                                "linha_excel": linha_excel,
                                "status": "erro_api",
                                "erro": erro_msg,
                            }
                        )
                except Exception as e:
                    erro_api += 1
                    log_cb(f"[ERRO] Linha {linha_excel}: falha de requisicao - {e}")
                    detalhes.append(
                        {
                            "linha_excel": linha_excel,
                            "status": "erro_api",
                            "erro": f"falha de requisicao - {e}",
                        }
                    )

                progresso_cb(index + 1, total)

            return {
                "sucesso": sucesso,
                "erro_api": erro_api,
                "ignorados": ignorados,
                "resumo": f"Finalizado: {sucesso} criados, {erro_api} erros de API, {ignorados} ignorados.",
                "detalhes": detalhes,
            }
        finally:
            try:
                self.cliente.finalizar_sessao(headers)
                log_cb("[OK] Sessao GLPI finalizada.")
            except Exception as e:
                log_cb(f"[AVISO] Nao foi possivel finalizar a sessao: {e}")

    def preparar_planilha_fechamento(self, caminho):
        df = normalizar_colunas(self.ler_planilha(caminho))
        if "ticket_id" not in df.columns:
            raise ValueError("A planilha deve conter a coluna obrigatoria: ticket_id")
        return df

    def preparar_planilha_solucao(self, caminho):
        df = normalizar_colunas(self.ler_planilha(caminho))
        colunas_obrigatorias = {"ticket_id", "solucao"}
        faltantes = [col for col in colunas_obrigatorias if col not in df.columns]
        if faltantes:
            raise ValueError(f"A planilha deve conter as colunas obrigatorias: {', '.join(sorted(faltantes))}")
        return df

    def fechar_chamados(self, df, usar_html, log_cb, progresso_cb):
        total = len(df)
        possui_solucao = "solucao" in df.columns
        headers = self.cliente.iniciar_sessao()
        log_cb("[OK] Sessao GLPI iniciada para fechamento.")

        fechados = 0
        erros = 0
        ignorados = 0

        try:
            for index, row in df.iterrows():
                linha_excel = index + 2
                ticket_id = int_or_none(row.get("ticket_id"))
                if ticket_id is None:
                    ignorados += 1
                    log_cb(f"[AVISO] Linha {linha_excel}: ticket_id invalido. Ignorada.")
                    progresso_cb(index + 1, total)
                    continue

                try:
                    if possui_solucao:
                        texto_solucao = row.get("solucao")
                        if not is_empty(texto_solucao):
                            conteudo_solucao = preparar_texto_glpi(texto_solucao, usar_html=usar_html)
                            r_solucao = self.cliente.adicionar_solucao(headers, ticket_id, conteudo_solucao)
                            if r_solucao.status_code not in (200, 201):
                                log_cb(
                                    f"[AVISO] Ticket {ticket_id}: nao foi possivel salvar solucao (status {r_solucao.status_code})."
                                )

                    r_fechar = self.cliente.fechar_chamado(headers, ticket_id)
                    if r_fechar.status_code in (200, 201):
                        fechados += 1
                    else:
                        erros += 1
                        log_cb(
                            f"[ERRO] Linha {linha_excel} / Ticket {ticket_id}: status {r_fechar.status_code} - {r_fechar.text[:180]}"
                        )
                except Exception as e:
                    erros += 1
                    log_cb(f"[ERRO] Linha {linha_excel} / Ticket {ticket_id}: falha de requisicao - {e}")

                progresso_cb(index + 1, total)

            return {
                "fechados": fechados,
                "erros": erros,
                "ignorados": ignorados,
                "resumo": f"Fechamento finalizado: {fechados} fechados, {erros} erros, {ignorados} ignorados.",
            }
        finally:
            try:
                self.cliente.finalizar_sessao(headers)
                log_cb("[OK] Sessao GLPI finalizada.")
            except Exception as e:
                log_cb(f"[AVISO] Nao foi possivel finalizar a sessao: {e}")

    def solucionar_chamados(self, df, usar_html, log_cb, progresso_cb):
        total = len(df)
        headers = self.cliente.iniciar_sessao()
        log_cb("[OK] Sessao GLPI iniciada para inclusao de solucao.")

        solucionados = 0
        erros = 0
        ignorados = 0

        try:
            for index, row in df.iterrows():
                linha_excel = index + 2
                ticket_id = int_or_none(row.get("ticket_id"))
                texto_solucao = row.get("solucao")

                if ticket_id is None:
                    ignorados += 1
                    log_cb(f"[AVISO] Linha {linha_excel}: ticket_id invalido. Ignorada.")
                    progresso_cb(index + 1, total)
                    continue

                if is_empty(texto_solucao):
                    ignorados += 1
                    log_cb(f"[AVISO] Linha {linha_excel} / Ticket {ticket_id}: solucao vazia. Ignorada.")
                    progresso_cb(index + 1, total)
                    continue

                try:
                    conteudo_solucao = preparar_texto_glpi(texto_solucao, usar_html=usar_html)
                    r_solucao = self.cliente.adicionar_solucao(headers, ticket_id, conteudo_solucao)
                    if r_solucao.status_code in (200, 201):
                        solucionados += 1
                    else:
                        erros += 1
                        log_cb(
                            f"[ERRO] Linha {linha_excel} / Ticket {ticket_id}: status {r_solucao.status_code} - {r_solucao.text[:180]}"
                        )
                except Exception as e:
                    erros += 1
                    log_cb(f"[ERRO] Linha {linha_excel} / Ticket {ticket_id}: falha de requisicao - {e}")

                progresso_cb(index + 1, total)

            return {
                "solucionados": solucionados,
                "erros": erros,
                "ignorados": ignorados,
                "resumo": f"Solucoes finalizadas: {solucionados} adicionadas, {erros} erros, {ignorados} ignorados.",
            }
        finally:
            try:
                self.cliente.finalizar_sessao(headers)
                log_cb("[OK] Sessao GLPI finalizada.")
            except Exception as e:
                log_cb(f"[AVISO] Nao foi possivel finalizar a sessao: {e}")

    def salvar_relatorio_importacao(self, resultado, log_texto):
        pasta_relatorios = Path.home() / "Documents" / "logchamados"
        try:
            pasta_relatorios.mkdir(parents=True, exist_ok=True)
        except Exception:
            pasta_relatorios = self.config_path.parent / "logchamados"
            pasta_relatorios.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_nome = f"importacao_{timestamp}"
        caminho_xlsx = pasta_relatorios / f"{base_nome}_tickets_criados.xlsx"
        caminho_log = pasta_relatorios / f"{base_nome}.log.txt"

        tickets_criados = []
        for item in resultado.get("detalhes", []):
            if item.get("status") == "criado" and item.get("ticket_id") is not None:
                tickets_criados.append({"ticket_id": item.get("ticket_id")})

        pd.DataFrame(tickets_criados, columns=["ticket_id"]).to_excel(caminho_xlsx, index=False)

        with caminho_log.open("w", encoding="utf-8") as f:
            f.write(log_texto)

        return str(caminho_xlsx), str(caminho_log)
