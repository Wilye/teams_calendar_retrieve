import requests
import json
from personal_config import *
from datetime import datetime
from personal_utils import start_and_end_of_next_week

start_of_next_week = start_and_end_of_next_week()[0]
end_of_next_week = start_and_end_of_next_week()[1]   

def get_channel_messages(access_token):
    graph_url = f"https://graph.microsoft.com/v1.0/teams/{group_id}/channels/{channel_id}/messages/"
    
    all_messages = []
    while True:
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Prefer': 'outlook.timezone="Central Standard Time"'
        }   
        response = requests.get(graph_url, headers=headers)
        
        response.raise_for_status()
        data = response.json()
        all_messages.extend(data['value'])
        if '@odata.nextLink' in data:
            graph_url = data['@odata.nextLink']
        else:
            break

    return all_messages

def get_exchange_ids(messages, access_token):
    exchange_ids = []
    for message in messages:
        for attachment in message.get('attachments', []):
            if attachment.get('contentType') == 'meetingReference':
                exchange_id = json.loads(attachment.get('content', "{}")).get("exchangeId")
                if exchange_id:
                    # Fetch the referenced event
                    graph_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/events/{exchange_id}"

                    event = get_event_or_none(graph_url, access_token)
                    if event is not None:
                        if is_within_next_week(event):
                            exchange_ids.append(event)
    
    return exchange_ids

def is_within_next_week(event):
    # If the event starts during next week, keep the message
    event_start_datetime_str = event['start']['dateTime']
    event_start_datetime_str = event_start_datetime_str[:-1] if len(event_start_datetime_str.split('.')[-1]) > 6 else event_start_datetime_str

    event_end_datetime_str = event['end']['dateTime']
    event_end_datetime_str = event_end_datetime_str[:-1] if len(event_end_datetime_str.split('.')[-1]) > 6 else event_end_datetime_str

    event_start = datetime.strptime(event_start_datetime_str, "%Y-%m-%dT%H:%M:%S.%f")
    event_end = datetime.strptime(event_end_datetime_str, "%Y-%m-%dT%H:%M:%S.%f")

    if not(event_end < start_of_next_week or event_start > end_of_next_week):
        return True
    else:
        return False
    
def get_calendar_events(exchange_ids, access_token):
    calendar_events = []
    for exchange_id in exchange_ids:
        graph_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/events/{exchange_id}"
        event = get_event_or_none(graph_url, access_token)
        if event is not None:
            calendar_events.append(event)
    return calendar_events


def get_event_or_none(graph_url, access_token):
    try:
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Prefer': 'outlook.timezone="Central Standard Time"'
        }  
        response = requests.get(graph_url, headers=headers)     
        response.raise_for_status()
        return response.json()
    except requests.exceptions.HTTPError as err:
        if err.response.status_code == 404:
            # This is a cancelled meeting, skip it
            print(f"Meeting at {graph_url} was cancelled.")
            return None
        else:
            # Some other error occurred, re-raise the exception
            raise