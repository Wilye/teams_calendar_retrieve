from datetime import datetime, time, timedelta
from personal_utils import *
import pandas as pd

# The function to process the events
def process_events(events):
    data = []
    for event in events:
        # Convert the start and end times to datetime objects
        start = datetime.fromisoformat(event['start']['dateTime'].split('.')[0]) 
        end = datetime.fromisoformat(event['end']['dateTime'].split('.')[0])  

        # Correct the end time if the event ends exactly at midnight
        if end.time() == time(0, 0, 0):
            end -= timedelta(seconds=1)

        name = next((attendee['emailAddress']['name'] for attendee in event['attendees'] if attendee['emailAddress']['name'] != "JERA_Americas_IT"), None) # Because organizer is the group Teams email so JERA_Amercias_IT teams in this case
        
        start_date = start.date()
        end_date = end.date()
        start_time = start.time()
        end_time = end.time()
        
        duration = "All day" if event['isAllDay'] else f"{start_time.strftime('%I:%M %p')} - {end_time.strftime('%I:%M %p')}"
        dates = start_date.strftime("%m-%d-%Y") if start_date == end_date else f"{start_date.strftime('%m-%d-%Y')} - {end_date.strftime('%m-%d-%Y')}"
        subject_title = event["subject"]
        
        # Append the data with the start date for sorting purposes
        data.append([name, start_date, subject_title, dates, duration])

    # Sort data by last name then first name and start date
    data.sort(key=lambda x: (get_names(x[0]), x[1]))

    # Create DataFrame without the start date column
    df = pd.DataFrame(data, columns=['Name', 'Start Date', 'Subject', 'Date(s)', 'Duration (CTZ)'])
    df = df.drop(columns=['Start Date'])

    return df

def df_to_html(df):
    # Convert DataFrame to HTML without index column
    if df.empty:
        events_html = '<p>No one is taking PTO or traveling for the next week</p>'
        events_title = ''
    else:
        events_html = df.to_html(index=False)
        events_title = '<h2 style="text-align: center;">PTO-Travel Schedule</h2>'

    return events_title, events_html