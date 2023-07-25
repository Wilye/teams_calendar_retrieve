import requests
from JERA_config import group_id
from JERA_utils import start_and_end_of_next_week
from datetime import timedelta

start_of_next_week = start_and_end_of_next_week()[0]
end_of_next_week = start_and_end_of_next_week()[1]

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

start_date = start_of_next_week.strftime("%m/%d/%Y")
end_date = end_of_next_week.strftime("%m/%d/%Y")
next_start_date = (start_of_next_week + timedelta(weeks=1)).strftime("%m/%d/%Y")
next_end_date = (end_of_next_week + timedelta(weeks=1)).strftime("%m/%d/%Y")

subject = f"{start_date} - {end_date} PTO and Travel Schedule Information"

body = f"""
<p>Hello ICT department,</p>

<p>This is a friendly reminder that you are expected to log your PTO and travel schedule to share with this team in the PTO-Travel channel on Microsoft Teams in the JERA Americas_IT Teams in advance.</p>

<p>By the end of today Friday, please make sure that you have logged your PTO and travel time that you will begin to take from {next_start_date} - {next_end_date} on the <a href="https://teams.microsoft.com/l/channel/19%3a487f4479c76746ceaae6b24b89c7a2c4%40thread.tacv2/PTO-%2520Travel?groupId=9717b8b0-c079-4946-a99b-3cae873d321e&tenantId=2840389b-0f81-496f-b742-ac794a5da61e">PTO-Travel channel calendar</a>.</p>

<p>Thank you very much for your contributions.</p>

<p>The list below is the recommended or required other action item to prepare for your time off:</p>

<ul>
    <li>Request time off from TimeClock.</li>
    <li>Cancel meetings that you organize or let the team know you can’t attend if needed.</li>
    <li>Set automatic replies in outlook if needed.</li>
    <li>Block your calendar by creating an all-day appointment with OOO mode so you don’t get new meeting invites during the time if needed.</li>
</ul>

<p>Here is the list of people taking time off and/or travelling in the upcoming week ({start_date} - {end_date}):</p>
"""

def get_teams_members_emails(access_token):
    teams_members_graph_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/members"
    headers = {
            'Authorization': f'Bearer {access_token}',
            'Prefer': 'outlook.timezone="Central Standard Time"'
    }  
    response_teams_members = requests.get(teams_members_graph_url, headers=headers) 
    members = response_teams_members.json()['value']
    
    emails = []
    for member in members:
        email = member.get("mail")
        if email:  # If the email is not None or empty
            emails.append(email)

    return emails

def send_email(access_token, recipients, subject, html_body):
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    headers = {
        'Authorization' : 'Bearer ' + access_token,
        'Content-Type'  : 'application/json'
    }

    # TEST: EMAIL BODY
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
                    "address": "Shelby.Yang_intern@jeraamericas.com"
                }
            }
        ]
    },
    "saveToSentItems": "true"
    }
    
    """
    email_body = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html_body
            },
            "toRecipients": recipients
        },
        "saveToSentItems": "true"
    }
    """

    # Send the email
    response = requests.post(url, headers=headers, json=email_body)

    # Check the response
    if response.status_code == 202:
        print("Email sent successfully!")
    else:
        print(f"Email not sent. Status code: {response.status_code}, Error: {response.text}")

def get_recipients(access_token):
    emails = get_teams_members_emails(access_token)
    recipients = [{"emailAddress": {"address": email}} for email in emails]
    return recipients