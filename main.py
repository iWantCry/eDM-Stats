import win32com.client as win32
import datetime

win32c = win32.constants

from askDir import askDir
from edm import edm
from convertFiletype import convertFiletype

def main():

    # Get current date
    year = datetime.datetime.today().year
    month = datetime.datetime.today().month
    day = datetime.datetime.today().day

    # Add leading zero to month
    if len(str(month)) == 1:
        month = f'0{month}'

    if len(str(day)) == 1:
        day = f'0{day}'

    # function calls
    path = askDir()

    path = convertFiletype(path, year, month, day)

    edm(path)

    print("File Tagging Complete!")
main()
