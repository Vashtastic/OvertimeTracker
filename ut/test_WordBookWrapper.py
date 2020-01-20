import pytest
from openpyxl import Workbook
from openpyxl import load_workbook
from Config import EMPLOYEE_NAME
from Config import EXCEL_FILE_NAME
from Config import SHEET
import datetime
import math
import enum

def test_awef():
    x = 1
    y = 2
    expected_sum = x + y
    assert 3 is expected_sum