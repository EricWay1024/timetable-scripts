from ics import Calendar, Event
from datetime import date, timedelta
import pandas as pd

calendar = Calendar()

first_sunday = date(2022, 9, 11)

day_of_week = {
    'Monday': 1,
    'Tuesday': 2,
    'Wednesday': 3,
    'Thursday': 4,
    'Friday': 5,
}

def parse_week_str(week_str: str):
    def helper(week_str_list: list):
        if not week_str_list: return []
        l = list(map(int, week_str_list[0].split('-')))
        if len(l) == 1:
            a, = l
            b, = l
        else:
            a, b = l
        return list(range(a, b + 1)) + helper(week_str_list[1:])
    return sorted(helper(week_str.replace(' ', '').split(',')))

def standardize_time(time_str):
    if len(time_str) == 4:
        return '0' + time_str
    else:
        return time_str

def remove_z(ics_line: str):
    if ics_line.startswith('DTEND') or ics_line.startswith('DTSTART'):
        return ics_line.replace('Z', '')
    else:
        return ics_line

def generate_events(weeks, day, start_time, end_time, location, title, type, professor, code, notes):
    for week in weeks:
        event = Event()
        event.name = f'{title} - {type}'
        event.location = location
        date_str = (first_sunday + timedelta(days=day + 7 * week)).isoformat()
        event.begin = f'{date_str} {start_time}'
        event.end = f'{date_str} {end_time}'
        event.description = f'{code} {professor} {notes}'
        calendar.events.add(event)


def generate_calendar_from_excel(excel_file_path, ics_file_path):
    df = pd.read_excel(excel_file_path)
    for _, r in df.iterrows():
        code = r[0]
        title = r[1]
        type = r[3]
        weeks = parse_week_str(r[4])
        day = day_of_week[r[5]]
        start_time = standardize_time(r[6])
        end_time = standardize_time(r[7])
        professor = r[8]
        location = r[9]
        notes = r[10]

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
            notes=notes
        )

    with open(ics_file_path, 'w') as my_file:
        my_file.writelines(map(remove_z, calendar.serialize_iter()))
    
if __name__ == '__main__':
    generate_calendar_from_excel('sample.xlsx', 'sample.ics')
