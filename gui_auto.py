import pywinauto as pw
from pywinauto import application
import os
import time

def export_dxf(path):

    draftsight_path = r"C:\Program Files\Dassault Systemes\DraftSight\bin\DraftSight.exe"
    app = application.Application().start(draftsight_path + " " + path)
    draftsight = app.window(title_re=".*DraftSight*")
    time.sleep(5)
    draftsight.type_keys("_SAVEAS{ENTER}")
    time.sleep(1)
    save_as = app.window(title_re=".*Save As*")
    save_as.type_keys("{TAB}")
    save_as.type_keys("{DOWN 9}{ENTER 2}")
    time.sleep(2)

    #Find out if dxf already exists.
    path_list = path.replace('\"', '').split('\\')
    path_list = path_list[:len(path_list)-1]
    #Splat: https://stackoverflow.com/questions/2322355/proper-name-for-python-operator
    dir_list = os.listdir(os.path.join(*path_list))

    if any(".dxf" in s for s in dir_list):
        top_window = app.top_window_()
        time.sleep(1)
        top_window.type_keys("{TAB}{ENTER}")


    draftsight.type_keys("%{F4}")

def export_pdf(path, sheets_quantities):

    draftsight_path = r"C:\Program Files\Dassault Systemes\DraftSight\bin\DraftSight.exe"
    app = application.Application().start(draftsight_path + " " + path)
    draftsight = app.window(title_re=".*DraftSight*")
    time.sleep(7)
    draftsight.type_keys("_EXPORTPDF{ENTER}")
    time.sleep(2)
    export_pdf_window = app.window(title_re=".*PDF*")
    export_pdf_window.type_keys("{TAB 2}{SPACE}")

    #REVS, INDEX, NOTES SHEETS
    export_pdf_window.type_keys("{DOWN}{SPACE}")
    export_pdf_window.type_keys("{DOWN}{SPACE}")
    export_pdf_window.type_keys("{DOWN}{SPACE}")

    for sheet in range(sheets_quantities):
        export_pdf_window.type_keys("{DOWN}{SPACE}")

    export_pdf_window.type_keys("{TAB 4}")
    export_pdf_window.type_keys("{UP}")
    export_pdf_window.type_keys("{TAB}{SPACE}")

    #export_pdf_window.print_control_identifiers()
    export_pdf_window.QWidget6.Click()
    export_pdf_window.type_keys("{DOWN 3}{ENTER}")

    #SAVE It auto selects the OK option
    export_pdf_window.type_keys("{ENTER}")

    time.sleep(2)

    #Find out if PDF already exists.
    path_list = path.replace('\"', '').split('\\')
    path_list = path_list[:len(path_list)-1]
    #Splat: https://stackoverflow.com/questions/2322355/proper-name-for-python-operator
    dir_list = os.listdir(os.path.join(*path_list))

    if any(".pdf" in s for s in dir_list):
        top_window = app.top_window_()
        top_window.type_keys("{ENTER}")

    draftsight.Wait('enabled', timeout=30)

    draftsight.type_keys("_EXIT{ENTER}")
    time.sleep(2)
    top_window = app.top_window_()
    top_window.type_keys("{TAB}{ENTER}")

if __name__ == "__main__":

    path = os.path.join("T:",
                        "JOBS",
                        "2017",
                        "171033 Steamworks Brewing",
                        "171033 ELECTRICAL",
                        "171033 EL01.1164 REV A.dwg"
                        )

    path =  "\""+path+"\""

    export_dxf(path)
    33
    #export_pdf(path, 2)
