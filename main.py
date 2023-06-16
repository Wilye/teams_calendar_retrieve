import requests
import json
from datetime import datetime, time, timedelta
from msal import ConfidentialClientApplication, SerializableTokenCache
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
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
start_of_next_week = (datetime.now() + timedelta(days=-datetime.now().weekday(), weeks=1)).replace(hour=0, minute=0, second=0, microsecond=0)
end_of_next_week = (start_of_next_week + timedelta(days=4)).replace(hour=23, minute=59, second=59)
print(f"Start of next week: {start_of_next_week}")
print(f"End of next week: {end_of_next_week}")

# Retrieve calendar events from the Teams channel
group_id = 'd053caff-d520-4b1b-a294-2743fa694c8a'
graph_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/calendar/calendarView"
headers = {
    'Authorization': f'Bearer {access_token}',
    'Prefer': 'outlook.timezone="Central Standard Time"'
}
params = {
    'startDateTime': start_of_next_week.isoformat(),
    'endDateTime': end_of_next_week.isoformat()
}
print(f"Request parameters: {params}")

response = requests.get(graph_url, headers=headers, params=params)
response.raise_for_status()

# Parse the response
events = response.json()['value']
with open("events.json", "w") as outfile:
    json.dump(events, outfile, indent = 4)

# Filter the events and compile into tables
pto_events = [event for event in events if 'pto' in event['subject'].lower()]
travel_events = [event for event in events if ' at ' in event['subject'].lower()]

# Create lists to store the event data
pto_data = []
travel_data = []

def get_names(name):
    """Extract the first and last names from a full name and convert them to lowercase"""
    names = name.split()
    first_name = names[0].lower() if len(names) > 0 else ''
    last_name = names[-1].lower() if len(names) > 1 else ''
    return last_name, first_name

# The function to process the events
def process_events(events, isPTO):
    data = []
    for event in events:
        # Convert the start and end times to datetime objects
        start = datetime.fromisoformat(event['start']['dateTime'].split('.')[0]) 
        end = datetime.fromisoformat(event['end']['dateTime'].split('.')[0])  

        # Flag
        # Correct the end time if the event ends exactly at midnight
        if end.time() == time(0, 0, 0):
            end -= timedelta(seconds=1)

        print("start in process_events: " + start.strftime("%m/%d/%Y, %H:%M:%S"))
        print("end in process_events: " + end.strftime("%m/%d/%Y, %H:%M:%S"))

        # Only process the event if it overlaps with or is contained within next week
        if not (end < start_of_next_week or start > end_of_next_week):
            name = next((attendee['emailAddress']['name'] for attendee in event['attendees'] if attendee['emailAddress']['name'] != "test team"), None) # Need to replace test team with whatever group name JERA Americas_IT has
            
            start_date = max(start.date(), start_of_next_week.date()) #  updates the start_date to be the later of the event's start date and the start of next week
            end_date = min(end.date(), end_of_next_week.date()) # updates the end_date to be the earlier of the event's end date and the end of next week
            start_time = start.time()
            end_time = end.time()
            print("start: " + str(start_date))
            print("end: " + str(end_date))
            duration = "All day" if event['isAllDay'] else f"{start_time.strftime('%I:%M %p')} - {end_time.strftime('%I:%M %p')}"
            dates = str(start_date) if start_date == end_date else f"{start_date} - {end_date}"
            
            # Append the data with the start date for sorting purposes
            if(isPTO): 
                data.append([name, start_date, dates, duration])
            else:
                data.append([name, start_date, dates])

    # Sort data by last name then first name and start date
    data.sort(key=lambda x: (get_names(x[0]), x[1]))

    # Create DataFrame without the start date column
    if(isPTO):
        df = pd.DataFrame(data, columns=['Name', 'Start Date', 'Date(s)', 'Duration (CST)'])
        df = df.drop(columns=['Start Date'])
    else:
        df = pd.DataFrame(data, columns=['Name', 'Start Date', 'Date(s)'])
        df = df.drop(columns=['Start Date'])

    return df

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
if pto_df.empty:
    pto_html = '<p>No one is taking PTO for the next week</p>'
    pto_title = ''
else:
    pto_html = pto_df.to_html(index=False)
    pto_title = '<h2 style="text-align: center;">PTO Schedule</h2>'

if travel_df.empty:
    travel_html = '<p>No one is travelling for the next week</p>'
    travel_title = ''
else:
    travel_html = travel_df.to_html(index=False)
    travel_title = '<h2 style="text-align: center;">Travel Schedule</h2>'

# Define email parameters
sender = "shelby@yangyy.onmicrosoft.com"
receiver = "shelby@yangyy.onmicrosoft.com"
password = "5750Jason" # TODO: setup an environment variable so password is not displayed in source code
# Define the date range for the subject
start_date = start_of_next_week.strftime("%m/%d/%Y")
end_date = end_of_next_week.strftime("%m/%d/%Y")
subject = f"{start_date} - {end_date} PTO and Travel Schedule Information"

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
combined_html = style + body + pto_title + pto_html + travel_title + travel_html
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

# Notes
"""
Need to test:
- make sure names are sorted by alphabetical last name, then first name, then date
    - so need to create a user with same last name, different first name
    - make a user with a Z last name to see if they appear behind me
- need to make sure that not all events with the word ' at ' are flagged as a Travel event because it's definitely possible that like an event in a different calendar has ' at ' in it's subject name
    - so perhaps i can bring up a way of consistent formatting like prefixing each event subject with [PTO] or [Travel]
        - so e.g. "[PTO] Shelby Yang" or "[Travel] Shelby at Houston"
        - this way it's less likely an event that's not PTO or Travel, but has ' at ' or 'xyzPTOxyz' in the subject will not be flagged
- if an event ends exactly at midnight, check to see if a second gets subtracted and it displays previous day
    - the thing is, if a person wants to create an event that DOES end at 12 AM on the dot, I think it might subtract a second, and so their date no longer reflects that day but the day before

Need to implement:
- running the script weekly (Fridays at 8 AM?)
- a no reply account with admin privileges (or get someone to grant the permissions i need) that will send out the email to the IT department
- more security, so sensitive information like password isn't just in the source code

Need to figure out:
- if token refreshing actually means the program should never need manual authorization (since the program will run on a weekly basis, the token will be refreshed on a weekly basis? or is it after the refresh token's time runs out, manual authorization is needed again)
    - https://stackoverflow.com/questions/51332122/access-token-refresh-token-with-msal this seems to answer
- the list of lowest permissions needed for the program to run

Useful shortcut: ^C to kill a process in the terminal
"""