import requests
import json
from datetime import datetime, timedelta
from msal import ConfidentialClientApplication
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


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
with open("events.json", "w") as outfile:
    json.dump(events, outfile, indent = 4)

# Filter the events and compile into tables
pto_events = [event for event in events if 'pto' in event['subject'].lower()]
travel_events = [event for event in events if 'at' in event['subject'].lower()]

# Create lists to store the event data
pto_data = []
travel_data = []

# The function to process the events
def process_events(events, isPTO):
    data = []
    for event in events:
        name = next((attendee['emailAddress']['name'] for attendee in event['attendees'] if attendee['emailAddress']['name'] != "test team"), None) # Need to replace test team with whatever group name JERA Americas_IT has
        start = event['start']['dateTime'].split("T")[0]
        end = event['end']['dateTime'].split("T")[0]
        duration = "All day" if event['isAllDay'] else f"{start} - {end}"
        if(isPTO): 
            data.append([name, f"{start} - {end}", duration])
        else:
            data.append([name, f"{start} - {end}"])
    if(isPTO):
        return pd.DataFrame(data, columns=['Name', 'Date(s)', 'Duration (CST)'])
    else:
        return pd.DataFrame(data, columns=['Name', 'Date(s)'])

# Parse the PTO events
pto_df = process_events(pto_events, True)

# Parse the travel events
# Add location parsing logic into process_events if it's needed for all events, or handle it separately like this:
travel_data = []
for event in travel_events:
    location = event['subject'].split('at')[1].strip()  # Assuming location follows 'travel' in the subject
    travel_data.append([event, location])

travel_df = process_events(travel_events, False)
travel_df['Location'] = [data[1] for data in travel_data]  # Add the parsed locations

# Convert DataFrames to HTML without index column
pto_html = pto_df.to_html(index=False)
travel_html = travel_df.to_html(index=False)

# Define email parameters
sender = "shelby@yangyy.onmicrosoft.com"
receiver = "shelby@yangyy.onmicrosoft.com"
password = "5750Jason"
subject = "Weekly PTO and Travel Events Report"

# Create email message
msg = MIMEMultipart()
msg['From'] = sender
msg['To'] = receiver
msg['Subject'] = subject

# HTML styling
style = """
<style>
    table {
        border-collapse: collapse;
        width: 100%;
    }
    th {
        text-align: center;
    }
</style>
"""

# Table titles
pto_title = "<h2>PTO Schedule</h2>"
travel_title = "<h2>Travel Schedule</h2>"

# Attach the tables to your email
combined_html = style + pto_title + pto_html + travel_title + travel_html
msg.attach(MIMEText(combined_html, 'html'))

# Writing to HTML file locally for debugging purposes
pto_html_local = open("pto_html","w")
travel_html_local = open("travel_html","w")

pto_html_local.write(pto_html)
travel_html_local.write(travel_html)

# Send the email
with smtplib.SMTP('smtp.outlook.com', 587) as server:
    server.starttls()
    server.login(sender, password)
    server.send_message(msg)

print('Email sent successfully!')