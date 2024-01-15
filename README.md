# Excel sheet of timetable → ICS file

## Usage

1. Edit `sample.xlsx`. University of Nottingham students can copy from [Teaching Timetables](https://timetabling.nottingham.ac.uk/) (select 'Autumn Semester' / 'Spring Semester' and 'List View'). Make sure each row has the same format as the sample rows. Delete sample rows, empty rows, rows containing only the day of week, and redundant headers. Tip: please make sure all entries are in the **'text'** format (in Microsoft Excel, select all cells => right click => format cells => select "text" from categories).
2. Make sure `first_sunday` defined in `main.py` is the same as the last Sunday before the first week of the semester; also change `local_tz` if needed.
3. Run `pip install pandas openpyxl ics && python3 main.py` and get your `sample.ics` file.
4. Upload the ICS file to your calendar. Tip: you might want to create a separate calendar and import to it, so that in case you find anything wrong you can delete the whole calendar and start all over again.

## Column Formats

- `Weeks`: Comma separated week ranges and weeks, where a week range from `a` to `b` inclusive is represented by `a-b`. E.g. `1-4,6,8` means weeks `1,2,3,4,6,8`.
- `Day`: Monday, Tuesday, Wednesday, Thursday, Friday.
- `Start`, `End`: 24-h time with format `HH:MM` or `H:MM`.
- Other columns are strings/texts.

## Unimportant Remarks

I worked for a long time on timetabling services (in particular, a web-app named uCourse) when I was at University of Nottingham Ningbo China. The idea of uCourse was to use a web scraper to fetch individual timetables from a university service so that each student could get their own ICS file with just one click. However, in practice, the situation can often be as messy as it gets, as people might change their courses during the term, and course information might also undergo updates. There is nothing more frustrating than running to a classroom according to your timetable, only to find no one there because the actual venue has been changed but not updated in the timetable cache when you exported the ICS file, and ironically this once happened to myself. 

At the end of the day, I realised that if you would spend time taking the actual courses, then it should not be a problem for you to spend less than twenty minutes editing an Excel spreadsheet containing the class information yourself. As long as this 'excel2ics' programme handles weekly repetition of classes well, it should be efficient enough. This also becomes a universal (category-theoretically speaking) script suitable for any student taking weekly repetitive classes. ~~Even better, you can't blame anyone else when the information is wrong because you typed it yourself.~~

I definitely considered turning this project into a web app, but it seemed to me that this would require implementing an online spreadsheet editing interface (much akin to Google Sheets), which is an unnecessary effort as you can edit things conveniently enough on Microsoft Excel or Google Sheets. I also (perhas wrongly) assume that most people already have Python installed on their machine these days, so this is it.