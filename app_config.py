import os
from dotenv import load_dotenv

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_PATH = "/getAToken"
SCOPE = ["User.Read", "People.Read", "Mail.Read", "Mail.Read.Shared", "Group.Read.All"]
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"

SECRET_KEY = os.getenv("FLASK_SECRET_KEY", "dev-secret-key-change-me")
SESSION_TYPE = "filesystem"
