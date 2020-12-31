from typing import Dict, Tuple

from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet


class StudentCollege:
    def __init__(self, name):
        self.name = name
        self.student_amount = 0

    def add_student(self):
        self.student_amount += 1


class Clazz:
    def __init__(self, name):
        self.name = name
        self.student_colleges: Dict[str, StudentCollege] = {}


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
        self.loser = row[18] != '正常'

    def cook_info(self, output):
        campus = output.setdefault(self.campus, Campus(self.campus))
        date = campus.dates.setdefault(self.date, ExamDate(self.date))
        session = date.sessions.setdefault(self.period, Session(self.period))
        course = session.courses.setdefault(self.course, Course(self.course, self.course_college))
        place = course.places.setdefault(self.place, Place(self.place))
        if self.loser:
            clazz = place.clazzes.setdefault('重修', Clazz('重修'))
        else:
            clazz = place.clazzes.setdefault(self.clazz, Clazz(self.clazz))

        student_college = clazz.student_colleges.setdefault(self.student_college, StudentCollege(self.student_college))
        student_college.add_student()


def process_file(input_filepath: str):
    wb = load_workbook(input_filepath)
    cooked_data = {}
    for name in wb.sheetnames:
        sheet: Worksheet = wb[name]
        for index, row in enumerate(sheet.values):
            if index == 0:
                continue
            try:
                row_info = RowInfo(row)
            except Exception:
                print(f"Can't parse row {index + 1}")
                continue
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

        sessions = list(date_obj.sessions.keys())
        sessions.sort()
        for session in sessions:
            session_obj = date_obj.sessions[session]

            for course, course_obj in session_obj.courses.items():
                places = list(course_obj.places.keys())
                places.sort()

                place_pointer = start_pointer
                student_college_pointer = start_pointer

                temp_student_college = None

                for place in places:
                    place_obj = course_obj.places[place]
                    for clazz_obj in place_obj.clazzes.values():
                        for student_college_obj in clazz_obj.student_colleges.values():
                            insert_row(sheet, end_pointer, date, session, course, course_obj.college, place,
                                       clazz_obj.name,
                                       student_college_obj.student_amount, student_college_obj.name)
                            if temp_student_college is None:
                                temp_student_college = student_college_obj.name
                            elif temp_student_college != clazz_obj.student_colleges:
                                sheet.merge_cells(start_column=8, end_column=8, start_row=student_college_pointer,
                                                  end_row=end_pointer - 1)
                                temp_student_college = student_college_obj.name
                                student_college_pointer = end_pointer

                            end_pointer += 1

                    if place_pointer != end_pointer - 1:
                        sheet.merge_cells(start_row=place_pointer, end_row=end_pointer - 1, start_column=5,
                                          end_column=5)
                    place_pointer = end_pointer

                sheet.merge_cells(start_row=start_pointer, end_row=end_pointer - 1, start_column=4, end_column=4)
                sheet.merge_cells(start_row=start_pointer, end_row=end_pointer - 1, start_column=3, end_column=3)
                sheet.merge_cells(start_row=start_pointer, end_row=end_pointer - 1, start_column=2, end_column=2)
                if student_college_pointer < end_pointer - 1:
                    sheet.merge_cells(start_row=student_college_pointer, end_row=end_pointer - 1, start_column=8,
                                      end_column=8)

                start_pointer = end_pointer


def save_file(data: Dict[str, Campus], output_filepath: str):
    wb = Workbook()
    wb.create_sheet("南湖校区", 0)
    wb.create_sheet("浑南校区", 1)

    if data.get("南湖校区"):
        save_to_sheet(data.get("南湖校区"), wb.worksheets[0])

    if data.get("浑南校区"):
        save_to_sheet(data.get("浑南校区"), wb.worksheets[1])

    wb.save(output_filepath)


data = process_file('Copy of 应考学生信息(2).xlsx')
save_file(data, './test.xlsx')
