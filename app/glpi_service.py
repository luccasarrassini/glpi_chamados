import requests


class GLPIClient:
    def __init__(self, api_url, app_token, user_token):
        self.api_url = api_url.strip().rstrip("/")
        self.app_token = app_token.strip()
        self.user_token = user_token.strip()

    def obter_headers_base(self):
        return {
            "App-Token": self.app_token,
            "Authorization": f"user_token {self.user_token}",
            "Content-Type": "application/json",
        }

    def autenticar(self):
        headers = self.obter_headers_base()
        r = requests.get(f"{self.api_url}/initSession", headers=headers, timeout=20)
        if r.status_code != 200:
            raise RuntimeError(f"status {r.status_code}: {r.text}")

        session_token = r.json().get("session_token")
        if not session_token:
            raise RuntimeError("resposta sem session_token.")

        headers["Session-Token"] = session_token
        try:
            requests.get(f"{self.api_url}/killSession", headers=headers, timeout=15)
        except Exception:
            pass

    def iniciar_sessao(self):
        headers = self.obter_headers_base()
        r = requests.get(f"{self.api_url}/initSession", headers=headers, timeout=20)
        if r.status_code != 200:
            raise RuntimeError(f"Falha initSession (status {r.status_code}): {r.text}")
        session_token = r.json().get("session_token")
        if not session_token:
            raise RuntimeError("initSession sem session_token na resposta.")
        headers["Session-Token"] = session_token
        return headers

    def finalizar_sessao(self, headers):
        requests.get(f"{self.api_url}/killSession", headers=headers, timeout=15)

    def criar_chamado(self, headers, payload):
        return requests.post(f"{self.api_url}/Ticket", headers=headers, json=payload, timeout=30)

    def usuario_existe(self, headers, user_id):
        r = requests.get(f"{self.api_url}/User/{user_id}", headers=headers, timeout=20)
        return r.status_code == 200

    def adicionar_solucao(self, headers, ticket_id, conteudo):
        payload = {
            "input": {
                "itemtype": "Ticket",
                "items_id": ticket_id,
                "content": conteudo,
            }
        }
        return requests.post(f"{self.api_url}/ITILSolution", headers=headers, json=payload, timeout=30)

    def fechar_chamado(self, headers, ticket_id):
        payload = {
            "input": {
                "id": ticket_id,
                "status": 6,
            }
        }
        return requests.put(f"{self.api_url}/Ticket/{ticket_id}", headers=headers, json=payload, timeout=30)

    def _buscar_item(self, headers, recurso, item_id):
        r = requests.get(f"{self.api_url}/{recurso}/{item_id}", headers=headers, timeout=20)
        if r.status_code != 200:
            return None
        try:
            return r.json()
        except Exception:
            return None

    @staticmethod
    def _extrair_nome_usuario(item):
        if not isinstance(item, dict):
            return ""
        nome = str(item.get("name", "")).strip()
        if nome:
            return nome
        realname = str(item.get("realname", "")).strip()
        firstname = str(item.get("firstname", "")).strip()
        completo = f"{firstname} {realname}".strip()
        return completo

    @staticmethod
    def _extrair_nome_generico(item):
        if not isinstance(item, dict):
            return ""
        for campo in ("completename", "name"):
            valor = str(item.get(campo, "")).strip()
            if valor:
                return valor
        return ""

    def obter_nome_usuario(self, headers, user_id):
        item = self._buscar_item(headers, "User", user_id)
        if not item:
            return None
        nome = self._extrair_nome_usuario(item)
        return nome or f"ID {user_id}"

    def obter_nome_categoria(self, headers, categoria_id):
        item = self._buscar_item(headers, "ITILCategory", categoria_id)
        if not item:
            return None
        nome = self._extrair_nome_generico(item)
        return nome or f"ID {categoria_id}"

    def obter_nome_localizacao(self, headers, localizacao_id):
        item = self._buscar_item(headers, "Location", localizacao_id)
        if not item:
            return None
        nome = self._extrair_nome_generico(item)
        return nome or f"ID {localizacao_id}"
