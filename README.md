# teams_calendar_retrieve_events

This is a project that sends the events in the upcoming week from a Microsoft Teams Calendar in an email.

Using Microsoft Graph API, it takes all messages in a Teams channel, filters through for those that are meeting events, looks up relevant information for each meeting (name, subject, start date, end date, duration), and compiles this information in a Pandas dataframe to be formatted as a table and sent in an email. The email list is automatically generated and consists of everyone in the Microsoft Teams.

I completed this project at my internship with JERA Americas.
