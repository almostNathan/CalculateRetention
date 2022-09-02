from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import tkinter
from tkinter import filedialog
import os


#get the workbook filepath from the user, dialog fileselect window
workbookFilePath = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select the Excel File", filetypes=(("Excel Files", "*.xlsx"),("All Files", "*.*")))
#open the workbook
workbook = load_workbook(workbookFilePath)
#get the active sheet
sheet = workbook.active
#get the maximum rows
maxRows = sheet.max_row

#get total dues collected
sheet.append(["Total Billed", f"=SUM(I2:I{maxRows})"])
sheet.append(["Total Billed", f"=SUM(IF(F2:F{maxRows}=FALSE,I2:I{maxRows},0))+SUM(IF(K2:K{maxRows}<>0,I2:I{maxRows},0))"])
sheet.append(["Retained",f"=B{maxRows+1}-B{maxRows+2}"])
workbook.save(workbookFilePath)
workbook.close