from openpyxl import Workbook
from openpyxl import load_workbook
from Config import EMPLOYEE_NAME
from Config import EXCEL_FILE_NAME
from Config import SHEET
import datetime
import math

HOUR_OFFSET = 12

class cell:
    def __init__(self, column_tag, row_tag):
        self.column_tag = column_tag
        self.row_tag = row_tag
    
    def __init__(self, cell):
        self.column_tag = cell[0]
        self.row_tag = cell[-1]
    
    def toString(self):
        return str(column_tag) + str(row_tag)

class HourFormat12:
    def __init__(self, hour, minute, meridian = "AM"):
        self.hour = hour - HOUR_OFFSET
        self.minute = minute
        self.meridian = "AM" if hour < 11 else "PM"

    def toString(self):
        return "{hour}:{minutes} {meridian}".format(hour=self.hour,
                                                  minutes=self.minute,
                                                  meridian=self.meridian)

class WorkbookWrapper:
    def __init__(self, WorkBookName, ExcelSheetFileName):
        print("Initializing Workbook for " + WorkBookName)
        self.excel_name = WorkBookName 
        self.work_book = load_workbook(WorkBookName)
        self.sheet_ranges = self.work_book[ExcelSheetFileName]
        self.start_ot_clock = "nan"
        self.finish_ot_clock = "nan"

    def print_cell(self, cell):
        print(self.sheet_ranges[cell].value)
    
    def write_to_cell(self, cell, value):
        self.sheet_ranges[cell] = value

    def save_workbook(self):
        print("Saving....")
        self.work_book.save(self.excel_name)
        print("Changes saved!")

    def create_copy_from_template():
        print("To be implemented")


class TrackerMenu:
    def __init__(self):
        self.employee_name = EMPLOYEE_NAME 
        self.workbook = WorkbookWrapper(EXCEL_FILE_NAME, SHEET)

    def print_menu(self):
        print("Hi " + EMPLOYEE_NAME + " What do you want to do today?")
        print(str("-" * 20))
        print("1. Start OT")
        print("2. End OT")
        print("3. Review Ot")
        print(str("-" * 20))
        choice = input("Enter your choice here: ")
        if int(choice) is 1:
            self.__start_ot()
        elif int(choice) is 2:
            self.__finish_ot()
        elif int(choice) is 3:
            self.__review_ot()

    @classmethod
    def get_current_date(self):
        current_date = datetime.datetime.now()
        return current_date.strftime("%d") + "/" + current_date.strftime("%m") + "/" + current_date.strftime("%Y")
    
    def __start_ot(self):
        current_date = datetime.datetime.now()
        hour_in_12hr_format = current_date.hour
        minute = current_date.minute
        self.workbook.start_ot_clock = HourFormat12(hour_in_12hr_format, minute)
        print ("Starting OT at " + self.workbook.start_ot_clock.toString())

    def __finish_ot(self):
        finish_minute_time = datetime.datetime.now().minute
        hour_in_12hr_format = datetime.datetime.now().hour
        self.workbook.finish_ot_clock = HourFormat12(hour_in_12hr_format, finish_minute_time)
        print ("Finished OT at " + self.workbook.finish_ot_clock.toString())

    def change_line_manager(self):
        manager_cell = cell()

    def display_current_time(self):
        DateStringFormat = get_current_date()
        return DateStringFormat
    
    def __review_ot(self):
        print("To be implemented")

menu = TrackerMenu()
menu.print_menu()