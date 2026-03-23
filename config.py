import os
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env para a memória
load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TEAM_ID = os.getenv("TEAM_ID")
SHAREPOINT_HOSTNAME = os.getenv("SHAREPOINT_HOSTNAME")
SHAREPOINT_SITE_PATH = os.getenv("SHAREPOINT_SITE_PATH")
