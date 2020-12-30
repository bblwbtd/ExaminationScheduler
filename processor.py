from typing import Dict, Tuple

from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet


class Clazz:
    def __init__(self, name, student_college):
        self.name = name
        self.student_college = student_college
        self.student_amount = 0

    def add_student(self):
        self.student_amount += 1


class Place:
    def __init__(self, name=''):
        self.name = name
        self.clazzes: Dict[str, Clazz] = {}


class Course:
    def __init__(self, name, college):
        self.name = name
        self.college = college
        self.places: Dict[str, Place] = {}


class Session:
    def __init__(self, period):
        self.period = period
        self.courses: Dict[str, Course] = {}


class ExamDate:
    def __init__(self, date):
        self.date = date
        self.sessions: Dict[str, Session] = {}


class Campus:
    def __init__(self, name):
        self.name = name
        self.dates: Dict[str, ExamDate] = {}


def parse_date_and_time(date_and_time: str):
    return tuple(date_and_time.split(" "))


class RowInfo:
    def __init__(self, row: Tuple):
        self.student_college = row[3]
        self.clazz = row[5]
        self.course = row[8]
        self.course_college = row[10]
        self.date, start_time = parse_date_and_time(row[11])
        _, end_time = parse_date_and_time(row[12])
        self.period = f'{start_time}-{end_time}'
        self.campus = row[13]
        self.place = row[15]

    def cook_info(self, output):
        campus = output.setdefault(self.campus, Campus(self.campus))
        date = campus.dates.setdefault(self.date, ExamDate(self.date))
        session = date.sessions.setdefault(self.period, Session(self.period))
        course = session.courses.setdefault(self.course, Course(self.course, self.course_college))
        place = course.places.setdefault(self.place, Place(self.place))
        clazz = place.clazzes.setdefault(self.clazz, Clazz(self.clazz, self.student_college))
        clazz.add_student()


def process_file(input_filepath: str):
    wb = load_workbook(input_filepath)
    sheet: Worksheet = wb[wb.sheetnames[0]]
    cooked_data = {}
    temp_course_name = ''

    for index, row in enumerate(sheet.values):
        if index == 0:
            continue
        try:
            row_info = RowInfo(row)
        except Exception:
            print(f"Can't parse row {index + 1}")
            continue
        if row_info.course is not None or '':
            temp_course_name = row_info.course
        else:
            row_info.course = temp_course_name
            row_info.cook_info(cooked_data)

    return cooked_data


def add_first_row(sheet: Worksheet):
    headers = ('考试日期', '考试时间', '考试名称', '开课院系', '考试地点', '班级', '考生人数', '考生院系')
    for index, header in enumerate(headers):
        sheet.cell(1, index + 1, header)


def insert_row(sheet: Worksheet, row: int, date, period, course, course_college, place, clazz, student_amount,
               student_college):
    temp = (date, period, course, course_college, place, clazz, student_amount, student_college)
    for index, item in enumerate(temp):
        sheet.cell(row, index + 1, item)


def adjust_column_style(sheet: Worksheet):
    sheet.column_dimensions['A'].width = 15
    sheet.column_dimensions['B'].width = 15
    sheet.column_dimensions['C'].width = 30
    sheet.column_dimensions['D'].width = 30
    sheet.column_dimensions['E'].width = 30
    sheet.column_dimensions['F'].width = 15
    sheet.column_dimensions['H'].width = 30


def save_to_sheet(campus: Campus, sheet: Worksheet):
    add_first_row(sheet)
    adjust_column_style(sheet)

    start_pointer = 2
    end_pointer = 2

    dates = list(campus.dates.keys())
    dates.sort()

    for date in dates:
        date_obj = campus.dates[date]
        for session, session_obj in date_obj.sessions.items():
            for course, course_obj in session_obj.courses.items():
                for place, place_obj in course_obj.places.items():
                    for clazz_obj in place_obj.clazzes.values():
                        insert_row(sheet, end_pointer, date, session, course, course_obj.college, place, clazz_obj.name,
                                   clazz_obj.student_amount, clazz_obj.student_college)
                        end_pointer += 1
                sheet.merge_cells(start_row=start_pointer, end_row=end_pointer - 1, start_column=3, end_column=3)
                sheet.merge_cells(start_row=start_pointer, end_row=end_pointer - 1, start_column=2, end_column=2)
                start_pointer = end_pointer


def save_file(data: Dict[str, Campus], output_filepath: str):
    wb = Workbook()
    wb.create_sheet("南湖校区", 0)
    wb.create_sheet("浑南校区", 1)

    save_to_sheet(data["南湖校区"], wb.worksheets[0])
    save_to_sheet(data["浑南校区"], wb.worksheets[1])

    wb.save(output_filepath)

#
# data = process_file('./Copy of 期末考试周应考学生信息 full.xlsx')
# save_file(data, "./test_full.xlsx")
