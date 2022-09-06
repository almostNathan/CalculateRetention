from openpyxl import load_workbook
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
#sheet[f"A{maxRows+1}"] = "=SUM(IF($F$2:$F${maxRows}=FALSE,$I$2:$I${maxRows},0))+SUM(IF(K2:K{maxRows}<>0,I2:I{maxRows},0))"

#total Billed
sheet[f"A{maxRows+1}"] = "Total Billed"
sheet[f"B{maxRows+1}"] = f"=SUM(I2:I{maxRows})"
#writtend Off
sheet[f"A{maxRows+2}"] = "Written Off"
sheet[f"B{maxRows+2}"] = f"=SUM(IF($F$2:$F${maxRows}=FALSE,$I$2:$I${maxRows},0))+SUM(IF(K2:K{maxRows}<>0,I2:I{maxRows},0))"
sheet[f"C{maxRows+2}"] = f"=B{maxRows+2}/B{maxRows+1}"
sheet[f"C{maxRows+2}"].number_format = "0.00%"
#retained
sheet[f"A{maxRows+3}"] = "Retained "
sheet[f"B{maxRows+3}"] = f"=B{maxRows+1}-B{maxRows+2}"
sheet[f"C{maxRows+3}"] = f"=B{maxRows+3}/B{maxRows+1}"
sheet[f"C{maxRows+3}"].number_format = "0.00%"

#fix nested IF
sheet.formula_attributes[f"B{maxRows+2}"] = {'t':'array', 'ref': f'B{maxRows+2}'}


#sheet.append(["Retained",f"=B{maxRows+1}-B{maxRows+2}"])

workbook.save(workbookFilePath)
workbook.close

os.startfile(workbookFilePath)
