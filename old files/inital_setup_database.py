import os
import requests
import pyodbc
from msal import ConfidentialClientApplication, SerializableTokenCache

def get_token():
    TOKEN_CACHE_FILE = "token_cache.pkl"

    # your Azure AD registered app's id and secret
    client_id = '7d844cd4-d153-4b2b-a438-8839061e2035'
    client_secret = 'ZYS8Q~g7AnAm~.s1HJDI~4mYH5DFgaD5uZ-WRczJ'
    tenant_id = '5ce202cb-b98c-4a2c-b703-9495a1d48b51'

    # Create a serializable token cache
    token_cache = SerializableTokenCache()

    # Create a confidential client application
    app = ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
        token_cache=token_cache,
    )

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

        # Save the state of the updated token cach e to a file
        with open(TOKEN_CACHE_FILE, 'w') as f:
            f.write(token_cache.serialize())

    return result['access_token']


def get_messages(access_token, group_id, channel_id, filter_string=None):
    graph_url = f"https://graph.microsoft.com/v1.0/teams/{group_id}/channels/{channel_id}/messages"
    if filter_string is not None:
        graph_url += f"?$filter={filter_string}"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Prefer': 'outlook.timezone="Central Standard Time"'
    }

    messages = []
    while True:
        response = requests.get(graph_url, headers=headers)
        response.raise_for_status()
        data = response.json()
        messages.extend(data['value'])
        if '@odata.nextLink' in data:
            graph_url = data['@odata.nextLink']
        else:
            break

    return messages


def insert_messages_into_db(connection, messages):
    cursor = connection.cursor()
    for message in messages:
        MessageID = message['id']
        
        # Check if the 'from' field and 'user' field exists
        if message.get('from') and message['from'].get('user'):
            UserID = message['from']['user']['id']
        else:
            UserID = 'System'  # System means that this is an automatically generated message
        LastModifiedDateTime = message['lastModifiedDateTime']
        
        # Ensure message has attachments before trying to access them
        if message['attachments']:
            ContentType = message['attachments'][0]['contentType']
            Content = message['attachments'][0]['content']
            AttachmentName = message['attachments'][0]['name']
        else:
            ContentType = Content = AttachmentName = None 

        # Use the SQL Server MERGE statement to either insert or update the message in the database
        sql_merge = """
        MERGE TeamsMessages AS target
        USING (VALUES (?, ?, ?, ?, ?, ?)) AS source (MessageID, UserID, LastModifiedDateTime, ContentType, Content, AttachmentName)
        ON (target.MessageID = source.MessageID)
        WHEN MATCHED AND target.LastModifiedDateTime < source.LastModifiedDateTime THEN
            UPDATE SET UserID = source.UserID, LastModifiedDateTime = source.LastModifiedDateTime, ContentType = source.ContentType, Content = source.Content, AttachmentName = source.AttachmentName
        WHEN NOT MATCHED THEN
            INSERT (MessageID, UserID, LastModifiedDateTime, ContentType, Content, AttachmentName)
            VALUES (source.MessageID, source.UserID, source.LastModifiedDateTime, source.ContentType, source.Content, source.AttachmentName);
        """
        cursor.execute(sql_merge, (MessageID, UserID, LastModifiedDateTime, ContentType, Content, AttachmentName))

        connection.commit()


def main():
    # Configure SQL connection
    server = 'teams-messages.database.windows.net' 
    database = 'teams-messages' 
    username = 'shelbyyang' 
    password = 'SQLpassword7123' 
    driver= '{ODBC Driver 18 for SQL Server}'

    connection = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';PORT=1433;DATABASE='+database+';UID='+username+';PWD='+ password)

    # Acquire access token
    access_token = get_token()

    # Retrieve all messages from the Teams channel
    group_id = 'd053caff-d520-4b1b-a294-2743fa694c8a'
    channel_id = '19:1e0fh5m1C_1dHyjQh9Q-itLcgoA5cIe7JdJvc46hqOk1@thread.tacv2'

    messages = get_messages(access_token, group_id, channel_id)
    insert_messages_into_db(connection, messages)


if __name__ == '__main__':
    main()