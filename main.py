import requests
import json
from datetime import datetime, timedelta
from msal import ConfidentialClientApplication

# your Azure AD registered app's id and secret
client_id = '7d844cd4-d153-4b2b-a438-8839061e2035'
client_secret = 'ZYS8Q~g7AnAm~.s1HJDI~4mYH5DFgaD5uZ-WRczJ'
tenant_id = '5ce202cb-b98c-4a2c-b703-9495a1d48b51'

# Create a confidential client application
app = ConfidentialClientApplication(
    client_id,
    authority=f"https://login.microsoftonline.com/{tenant_id}",
    client_credential=client_secret,
)

# Redirect the user to the authorization URL
auth_url = app.get_authorization_request_url(["https://graph.microsoft.com/.default"])
print(f"Please go to this URL and authorize the app: {auth_url}")

# Get the authorization code from the user
auth_code = input("Enter the authorization code: ")

# Acquire a token using the authorization code
result = app.acquire_token_by_authorization_code(auth_code, ["https://graph.microsoft.com/.default"], redirect_uri="https://localhost/")
#print(result)
access_token = result['access_token']

# Calculate the start and end of next week
start_of_next_week = datetime.now() + timedelta(days=-datetime.now().weekday(), weeks=1)
end_of_next_week = start_of_next_week + timedelta(days=6)

# Retrieve calendar events from the Teams channel
group_id = 'd053caff-d520-4b1b-a294-2743fa694c8a'
graph_url = f'https://graph.microsoft.com/v1.0/groups/{group_id}/calendar/events'
headers = {
    'Authorization': f'Bearer {access_token}',
    #'Prefer': 'outlook.timezone="Central Standard Time"'
}
params = {
    'startDateTime': start_of_next_week.isoformat(),
    'endDateTime': end_of_next_week.isoformat()
}
response = requests.get(graph_url, headers=headers, params=params)
response.raise_for_status()

# Parse the response
events = response.json()['value']
with open("sample.json", "w") as outfile:
    json.dump(events, outfile, indent = 4)
#print(events)