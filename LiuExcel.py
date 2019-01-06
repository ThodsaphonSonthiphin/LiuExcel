import openpyxl
from openpyxl import Workbook


def loadWorkbook(name:str)->openpyxl.Workbook:

    return openpyxl.load_workbook(name)



