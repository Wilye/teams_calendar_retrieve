import requests
import json
from datetime import datetime, timedelta
from msal import PublicClientApplication

# your Azure AD registered app's id and secret
client_id = '7d844cd4-d153-4b2b-a438-8839061e2035'
client_secret = 'ZYS8Q~g7AnAm~.s1HJDI~4mYH5DFgaD5uZ-WRczJ'

# Get an access token  
token_url = 'https://login.microsoftonline.com/{5ce202cb-b98c-4a2c-b703-9495a1d48b51}/oauth2/v2.0/token'
token_data = {
    'grant_type': 'client_credentials',
    'client_id': client_id,
    'client_secret': client_secret,
    'scope': 'https://graph.microsoft.com/.default'
}

response = requests.post(token_url, data=token_data)
response.raise_for_status()
access_token = response.json()['access_token']

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