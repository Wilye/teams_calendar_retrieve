import requests
from datetime import datetime, timedelta
from msal import ConfidentialClientApplication, SerializableTokenCache
import os

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

    # Save the state of the updated token cache to a file
    with open(TOKEN_CACHE_FILE, 'w') as f:
        f.write(token_cache.serialize())

access_token = result['access_token']

# Calculate the start and end of next week
start_of_next_week = datetime.now() + timedelta(days=-datetime.now().weekday(), weeks=1)
start_of_next_week = start_of_next_week.replace(hour=0, minute=0, second=0, microsecond=0)


end_of_next_week = start_of_next_week + timedelta(days=6)
end_of_next_week = end_of_next_week.replace(hour=23, minute=59, second=59)

print(f"Start of next week: {start_of_next_week}")
print(f"End of next week: {end_of_next_week}")

# Define email parameters
sender = "shelby.yang_intern@jeraamericas.com"
receiver = "shelby.yang_intern@jeraamericas.com"
start_date = start_of_next_week.strftime("%m/%d/%Y")
end_date = end_of_next_week.strftime("%m/%d/%Y")
subject = f"{start_date} - {end_date} PTO and Travel Schedule Information"

# HTML styling
style = """
<style>
    table {
        border-collapse: collapse;
        width: 100%;
    }
    th, td, h2 {
        text-align: center;
    }
</style>
"""

# Define the date range for the body
next_start_date = (start_of_next_week + timedelta(weeks=1)).strftime("%m/%d/%Y")
next_end_date = (end_of_next_week + timedelta(weeks=1)).strftime("%m/%d/%Y")

# Format the body with these dates
body = f"""
<p>Hello ICT department,</p>

<p>If you haven't already, please remember to log your PTO and travel time in the PTO-Travel channel on Microsoft Teams in the JERA Americas_IT Teams at least one week ahead. By the end of today Friday, you should make sure that you have logged your PTO and travel time from {next_start_date} - {next_end_date} on the PTO-Travel channel calendar.</p>

<p>Here is the list of people taking time off and/or travelling in the upcoming week ({start_date} - {end_date}):</p>
"""

# Attach the tables to your email
combined_html = style + body

def send_email(access_token, email, subject, html_body):
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    headers = {
        'Authorization' : 'Bearer ' + access_token,
        'Content-Type'  : 'application/json'
    }

    # Prepare the email body
    email_body = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html_body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": email
                    }
                }
            ]
        },
        "saveToSentItems": "true"
    }

    # Send the email
    response = requests.post(url, headers=headers, json=email_body)

    # Check the response
    if response.status_code == 202:
        print("Email sent successfully!")
    else:
        print(f"Email not sent. Status code: {response.status_code}, Error: {response.text}")

# Call the send_email function
send_email(access_token, receiver, subject, combined_html)




    
    
    
    


