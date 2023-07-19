import json

from personal_auth import create_app, generate_access_token
from personal_data_processing import get_channel_messages, get_exchange_ids, get_calendar_events 
from personal_events_processing import process_events, df_to_html
from personal_email import get_recipients, style, body, send_email, subject

def main():
    app = create_app()
    access_token = generate_access_token(app)

    messages = get_channel_messages(access_token)
    with open("messages.txt", "w") as outfile:
        json.dump(messages, outfile, indent = 4)

    exchange_ids = get_exchange_ids(access_token, messages)
    #calendar_events = get_calendar_events(exchange_ids, access_token)

    #events_dataframe = process_events(calendar_events)
    #events_html = df_to_html(events_dataframe)

    #recipients = get_recipients(access_token)
    #combined_html = style + body + events_html[0] + events_html[1]
    
    #send_email(access_token, recipients, subject, combined_html)


if __name__ == "__main__":
    main()