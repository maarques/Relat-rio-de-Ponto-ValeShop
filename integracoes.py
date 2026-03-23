import requests
from config import *
import os

def get_access_token():
    print("Obtendo token de segurança...")
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "client_id": CLIENT_ID, "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET, "grant_type": "client_credentials"
    }
    response = requests.post(url, data=payload)
    return response.json().get("access_token") if response.status_code == 200 else None

def get_users_and_groups(token):
    print("Buscando perfis da equipe no Teams...")
    mapa = {}
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    res_members = requests.get(f"https://graph.microsoft.com/v1.0/groups/{TEAM_ID}/members", headers=headers)
    if res_members.status_code == 200:
        for m in res_members.json().get("value", []):
            email = m.get("userPrincipalName", "").lower()
            mapa[m.get("id")] = {"nome": m.get("displayName"), "email": email, "setor": "Sem Setor"}
            
    res_groups = requests.get(f"https://graph.microsoft.com/v1.0/teams/{TEAM_ID}/schedule/schedulingGroups", headers=headers)
    if res_groups.status_code == 200:
        for g in res_groups.json().get("value", []):
            setor = g.get("displayName")
            for uid in g.get("userIds", []):
                if uid in mapa: mapa[uid]["setor"] = setor
    return mapa

def encontrar_campo(dicionario, palavras_chave):
    for chave, valor in dicionario.items():
        for palavra in palavras_chave:
            if palavra.lower() in chave.lower() and valor is not None:
                return valor
    return None

def get_sharepoint_data(token, list_name, custom_site_path=None):
    # Se a gente não passar um site específico, ele usa o padrão do RH (.env)
    caminho_site = custom_site_path if custom_site_path else SHAREPOINT_SITE_PATH
    
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOSTNAME}:{caminho_site}:/lists/{list_name}/items?expand=fields"
    response = requests.get(url, headers=headers)
    return response.json().get("value", []) if response.status_code == 200 else []

def carregar_regras_rh(token):
    print("Lendo Livro de Regras (Contratos) no SharePoint...")
    contratos = {}
    itens = get_sharepoint_data(token, "Contratos Colaboradores")
    for item in itens:
        campos = item.get("fields", {})
        email = encontrar_campo(campos, ["email", "title", "título"])
        carga = encontrar_campo(campos, ["carga", "horaria"])
        almoco = encontrar_campo(campos, ["minuto", "almoco", "almoço"])
        
        if email and carga is not None and almoco is not None:
            email_limpo = str(email).lower().strip()
            contratos[email_limpo] = {"carga": float(carga), "almoco": int(almoco)}
    return contratos

def carregar_aprovacoes_he(token):
    print("Lendo Aprovações de Hora Extra no SharePoint...")
    aprovacoes = {}
    itens = get_sharepoint_data(token, "Aprovacoes Hora Extra")
    for item in itens:
        campos = item.get("fields", {})
        colaborador = encontrar_campo(campos, ["nome", "email", "colaborador"])
        data_he = encontrar_campo(campos, ["data"])
        status = encontrar_campo(campos, ["status", "resultado"])
        
        if colaborador and data_he:
            data_limpa = data_he.split("T")[0]
            chave = f"{str(colaborador).lower().strip()}_{data_limpa}"
            aprovacoes[chave] = status
    return aprovacoes

def obter_cartoes_ponto(token):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"https://graph.microsoft.com/v1.0/teams/{TEAM_ID}/schedule/timeCards"
    response = requests.get(url, headers=headers)
    return response.json().get("value", []) if response.status_code == 200 else []

def carregar_feriados(token):
    print("Lendo Calendário de Eventos (Feriados) na Intranet Principal...")
    feriados = set()
    
    # O Pulo do Gato: Dizendo pro robô ir no site correto!
    site_calendario = "/sites/IntranetValeShop"
    
    # Tentando ler a lista em português
    itens = get_sharepoint_data(token, "Eventos", custom_site_path=site_calendario) 
    
    # Se vier vazio, tenta em inglês
    if not itens:
        itens = get_sharepoint_data(token, "Events", custom_site_path=site_calendario)
        
    print(f"➜ Total de eventos encontrados no calendário: {len(itens)}")

    for item in itens:
        campos = item.get("fields", {})
        titulo = encontrar_campo(campos, ["title", "título"])
        data_evento = encontrar_campo(campos, ["eventdate", "start", "data de início"])
        categoria = encontrar_campo(campos, ["category", "categoria"])
        
        if titulo and data_evento:
            texto_busca = f"{str(titulo)} {str(categoria)}".lower()
            
            if any(palavra in texto_busca for palavra in ["feriado", "recesso", "folga"]):
                data_limpa = str(data_evento).split("T")[0]
                feriados.add(data_limpa)
                
    return feriados

def carregar_ferias(token):
    print("Lendo Aprovações de Férias no SharePoint (RH)...")
    ferias_aprovadas = {}
    
    itens = get_sharepoint_data(token, "Aprovações Férias")
    if not itens:
        itens = get_sharepoint_data(token, "AprovacoesFerias")
        
    for item in itens:
        campos = item.get("fields", {})
        
        email = encontrar_campo(campos, ["email", "title", "título"])
        # A correção que salvou o projeto ficou aqui:
        data_inicio = encontrar_campo(campos, ["inicio", "início", "start", "datadein_x00ed_cio"])
        data_fim = encontrar_campo(campos, ["fim", "end", "datadefim"])
        
        status_bruto = encontrar_campo(campos, ["status", "resultado"])
        status_texto = ""
        
        if isinstance(status_bruto, dict):
            status_texto = status_bruto.get("Value", "") or status_bruto.get("value", "")
        elif status_bruto:
            status_texto = str(status_bruto)
            
        if email and data_inicio and data_fim and status_texto:
            if status_texto.strip().upper() == "APROVADO":
                email_limpo = str(email).lower().strip()
                dt_ini = str(data_inicio).split("T")[0]
                dt_fim = str(data_fim).split("T")[0]
                
                if email_limpo not in ferias_aprovadas:
                    ferias_aprovadas[email_limpo] = []
                ferias_aprovadas[email_limpo].append((dt_ini, dt_fim))
                
    return ferias_aprovadas

def upload_excel_sharepoint(token, caminho_arquivo_local, nome_pasta_destino="Relatorios_Ponto"):
    print(f"\n🚀 Iniciando upload do arquivo para o SharePoint (RH)...")
    
    nome_arquivo = os.path.basename(caminho_arquivo_local)
    
    # PASSO 1: Pegar o ID real do site para evitar o bug de URL da Microsoft
    print("   -> Buscando ID do site...")
    site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOSTNAME}:{SHAREPOINT_SITE_PATH}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    
    site_response = requests.get(site_url, headers=headers)
    if site_response.status_code != 200:
        print(f"❌ Erro ao acessar o site: {site_response.json()}")
        return False
        
    site_id = site_response.json().get("id")
    
    # PASSO 2: Fazer o upload limpo usando o ID
    print(f"   -> Site encontrado! Enviando '{nome_arquivo}' para a nuvem...")
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{nome_pasta_destino}/{nome_arquivo}:/content"
    
    headers_upload = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
    
    try:
        with open(caminho_arquivo_local, "rb") as arquivo:
            conteudo_binario = arquivo.read()
            
        response = requests.put(upload_url, headers=headers_upload, data=conteudo_binario)
        
        if response.status_code in [200, 201]:
            print(f"✅ SUCESSO ABSOLUTO! Planilha enviada para a pasta '{nome_pasta_destino}'.")
            return True
        else:
            print(f"❌ Falha no upload: HTTP {response.status_code}")
            print(response.json())
            return False
            
    except FileNotFoundError:
        print(f"❌ Erro: O arquivo local não foi encontrado.")
        return False
    except Exception as e:
        print(f"❌ Erro inesperado ao enviar arquivo: {e}")
        return False
