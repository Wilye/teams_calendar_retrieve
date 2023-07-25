from datetime import datetime, timedelta

def start_and_end_of_next_week():
    start_of_next_week = datetime.now() + timedelta(days=-datetime.now().weekday(), weeks=1)
    start_of_next_week = start_of_next_week.replace(hour=0, minute=0, second=0, microsecond=0)

    end_of_next_week = start_of_next_week + timedelta(days=6)
    end_of_next_week = end_of_next_week.replace(hour=23, minute=59, second=59)

    return start_of_next_week, end_of_next_week

def get_names(name):
    names = name.split()
    first_name = names[0].lower() if len(names) > 0 else ''
    last_name = names[-1].lower() if len(names) > 1 else ''
    return last_name, first_name