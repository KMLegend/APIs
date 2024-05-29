import msal
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

# Replace with your own client ID and tenant ID
client_id=""
tenant_id=""

authority = f"https://tenant_name.ciamlogin.com"

app = msal.PublicClientApplication(
    client_id=client_id,
    authority=authority
)

# Acquire the token using the device code flow
result = app.acquire_token_by_device_flow(
    {"scopes": ["https://graph.microsoft.com/.default"]}
)

if "access_token" in result:
    token_de_acesso = result["access_token"]
else:
    print(result.get("error"))