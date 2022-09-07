from openpyxl import load_workbook
from tkinter import filedialog
import os


#get the workbook filepath from the user, dialog fileselect window
workbookFilePath = filedialog.askopenfilename(initialdir="C:/", title="Select the Excel File", filetypes=(("Excel Files", "*.xlsx"),("All Files", "*.*")))
#open the workbook
workbook = load_workbook(workbookFilePath)
#get the active sheet
sheet = workbook.active
#get the maximum rows
maxRows = sheet.max_row

#get total dues collected
#sheet[f"A{maxRows+1}"] = "=SUM(IF($F$2:$F${maxRows}=FALSE,$I$2:$I${maxRows},0))+SUM(IF(K2:K{maxRows}<>0,I2:I{maxRows},0))"

#written Off
sheet[f"A{maxRows+1}"] = "Written Off"
sheet[f"B{maxRows+1}"] = f"=SUMIFS(I2:I{maxRows},F2:F{maxRows},FALSE)"
#percent written off
sheet[f"C{maxRows+1}"] = f"=B{maxRows+1}/B{maxRows+4}"
sheet[f"C{maxRows+1}"].number_format = "0.00%"

#Retained
sheet[f"A{maxRows+2}"] = "Collected"
sheet[f"B{maxRows+2}"] = f"=SUMIFS(I2:I{maxRows},F2:F{maxRows},TRUE,K2:K{maxRows},\"=0\")"
#percent Retained(collected)
sheet[f"C{maxRows+2}"] = f"=B{maxRows+2}/B{maxRows+4}"
sheet[f"C{maxRows+2}"].number_format = "0.00%"

#AB Due
sheet[f"A{maxRows+3}"] = "In Collections"
sheet[f"B{maxRows+3}"] = f"=SUMIFS(I2:I{maxRows},F2:F{maxRows},TRUE,K2:K{maxRows},\"<>0\")"
#percent AB Due
sheet[f"C{maxRows+3}"] = f"=B{maxRows+3}/B{maxRows+4}"
sheet[f"C{maxRows+3}"].number_format = "0.00%"

#total Billed
sheet[f"A{maxRows+4}"] = "Total Billed"
sheet[f"B{maxRows+4}"] = f"=SUM(I2:I{maxRows})"


workbook.save(workbookFilePath)
workbook.close

os.startfile(workbookFilePath)
