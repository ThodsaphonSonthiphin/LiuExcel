import openpyxl
from openpyxl import Workbook


def loadWorkbook(name:str)->openpyxl.Workbook:
    print("Hello world")

    return openpyxl.load_workbook(name)



