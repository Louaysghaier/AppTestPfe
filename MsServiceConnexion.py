

import msal
import requests

# Define the necessary parameters
client_id = 'YOUR_CLIENT_ID'
client_secret = 'YOUR_CLIENT_SECRET'
tenant_id = 'YOUR_TENANT_ID'
authority = f'https://login.microsoftonline.com/{tenant_id}'
scope = ['https://graph.microsoft.com/.default']

# Create a confidential client application
app = msal.ConfidentialClientApplication(
    client_id,
    authority=authority,
    client_credential=client_secret,
)

# Acquire a token
result = app.acquire_token_for_client(scopes=scope)

if 'access_token' in result:
    access_token = result['access_token']
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    # Example request to get SharePoint sites
    response = requests.get('https://graph.microsoft.com/v1.0/sites', headers=headers)
    
    if response.status_code == 200:
        sites = response.json()
        print(sites)
    else:
        print(f"Error: {response.status_code}")
        print(response.json())
else:
    print("Error acquiring token:")
    print(result.get("error"))
    print(result.get("error_description"))
    print(result.get("correlation_id"))