from JERA_auth import create_app, generate_access_token
from JERA_data_processing import get_channel_messages, get_exchange_ids, get_calendar_events 
from JERA_events_processing import process_events, df_to_html
from JERA_email import get_recipients, style, body, send_email, subject

def main():
    app = create_app()
    access_token = generate_access_token(app)

    messages = get_channel_messages(access_token)
    exchange_ids = get_exchange_ids(messages, access_token)
    calendar_events = get_calendar_events(exchange_ids, access_token)

    events_dataframe = process_events(calendar_events)
    events_html = df_to_html(events_dataframe)

    recipients = get_recipients(access_token)
    combined_html = style + body + events_html[0] + events_html[1]
    
    send_email(access_token, recipients, subject, combined_html)

if __name__ == "__main__":
    main()