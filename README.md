# Excel sheet of timetable → ICS file

## Usage

1. Edit `sample.xlsx`. University of Nottingham students can copy from [Teaching Timetables](https://timetabling.nottingham.ac.uk/) (select 'Autumn Semester' / 'Spring Semester' and 'List View'). Make sure each row has the same format as the sample rows. Delete sample rows, empty rows, rows containing only the day of week, and redundant headers. 
2. Make sure `first_sunday` defined in `main.py` is the same as the last Sunday before the first week of the semester.
3. Run `pip install pandas ics && python3 main.py` and get your `sample.ics` file.
4. Upload the ICS file to your calendar.

## Column Formats

- `Weeks`: Comma separated week ranges and weeks, where a week range from `a` to `b` inclusive is represented by `a-b`.
- `Day`: Monday, Tuesday, Wednesday, Thursday, Friday.
- `Start`, `End`: 24-h time with format `HH:MM` or `H:MM`.
- Other columns are strings.

