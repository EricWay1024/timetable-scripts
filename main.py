from ics import Calendar, Event, DisplayAlarm
from datetime import date, timedelta
from datetime import datetime
import pandas as pd
import pytz


calendar = Calendar()

first_sunday = date(2023, 10, 8)
local_tz = pytz.timezone("Europe/London")

day_of_week = {
    'Monday': 1,
    'Tuesday': 2,
    'Wednesday': 3,
    'Thursday': 4,
    'Friday': 5,
}


def parse_week_str(week_str: str):
    def helper(week_str_list: list):
        if not week_str_list:
            return []
        week_tuple = tuple(map(int, week_str_list[0].split('-')))
        if len(week_tuple) == 1:
            a, = week_tuple
            b, = week_tuple
        else:
            a, b = week_tuple
        return list(range(a, b + 1)) + helper(week_str_list[1:])
    return sorted(helper(week_str.replace(' ', '').split(',')))


def standardize_time(time_str):
    if len(time_str) == 4:
        return '0' + time_str
    else:
        return time_str


def local_to_iso(date_str: str):
    # Parse the input string to a datetime object
    input_datetime = datetime.strptime(date_str, "%Y-%m-%d %H:%M")
    # Localize the datetime to the London timezone
    local_datetime = local_tz.localize(input_datetime, is_dst=None)
    # Convert the localized datetime to ISO format
    iso_date_str = local_datetime.isoformat()
    return iso_date_str


def generate_events(weeks, day, start_time, end_time, location, title, type,
                    professor, code, notes, session):
    for week in weeks:
        event = Event()
        event.name = f'{type} - {title}'
        event.location = location
        date_str = (first_sunday + timedelta(days=day +
                    7 * (week - 1))).isoformat()
        event.begin = local_to_iso(f'{date_str} {start_time}')
        event.end = local_to_iso(f'{date_str} {end_time}')
        event.description = f'{code} {professor} {notes} {session}'.replace(
            'nan', '')
        event.alarms = [DisplayAlarm(trigger=timedelta(minutes=-20))]
        calendar.events.add(event)


def generate_calendar_from_excel(excel_file_path, ics_file_path):
    df = pd.read_excel(excel_file_path, dtype='str')
    for _, r in df.iterrows():
        code = r.iloc[0]
        title = r.iloc[1]
        session = r.iloc[2]
        type = r.iloc[3]
        weeks = parse_week_str(r.iloc[4])
        day = day_of_week[r.iloc[5]]
        start_time = standardize_time(r.iloc[6])
        end_time = standardize_time(r.iloc[7])
        professor = r.iloc[8]
        location = r.iloc[9]
        notes = r.iloc[10]

        generate_events(
            code=code,
            title=title,
            type=type,
            weeks=weeks,
            day=day,
            start_time=start_time,
            end_time=end_time,
            professor=professor,
            location=location,
            notes=notes,
            session=session,
        )

    with open(ics_file_path, 'w') as my_file:
        my_file.writelines(calendar.serialize_iter())


if __name__ == '__main__':
    generate_calendar_from_excel('sample.xlsx', 'sample.ics')
