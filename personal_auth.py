import os
import json
from msal import ConfidentialClientApplication, SerializableTokenCache
from personal_config import *

TOKEN_CACHE_FILE = "token_cache.pkl"

# Create a serializable token cache
token_cache = SerializableTokenCache()

def create_app():
    # Create a confidential client application
    app = ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
        token_cache=token_cache,
    )
    
    return app

def generate_access_token(app):
    # Load the state of the token cache from file, if it exists
    if os.path.exists(TOKEN_CACHE_FILE):
        try:
            with open(TOKEN_CACHE_FILE, 'r') as f:
                token_cache.deserialize(f.read())
        except EOFError:
            print("Token cache file is empty. Will create a new one.")

    result = None
    accounts = app.get_accounts()

    if accounts:
        # If possible, look up tokens from the cache
        result = app.acquire_token_silent(["https://graph.microsoft.com/.default"], account=accounts[0])

    if not result:
        # If cache lookup failed, acquire a new token
        # Redirect the user to the authorization URL
        auth_url = app.get_authorization_request_url(["https://graph.microsoft.com/.default"])
        print(f"Please go to this URL and authorize the app: {auth_url}")

        # Get the authorization code from the user
        auth_code = input("Enter the authorization code: ")

        # Acquire a token using the authorization code
        result = app.acquire_token_by_authorization_code(auth_code, ["https://graph.microsoft.com/.default"], redirect_uri="https://localhost/")

        # Save the state of the updated token cache to a file
        with open(TOKEN_CACHE_FILE, 'w') as f:
            f.write(token_cache.serialize())

    access_token = result['access_token']
    return access_token

