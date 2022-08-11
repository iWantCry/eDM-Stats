import win32com.client as win32
import os

def convertFiletype(filepath, year, month, day):
    fname = filepath
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)

    wb.SaveAs(f"{year}{month}{day} edm stats.xlsx", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()

    for r,d,f in os.walk("C:\\"):
        for files in f:
            if files == f"{year}{month}{day} edm stats.xlsx":
                return os.path.join(r,files)