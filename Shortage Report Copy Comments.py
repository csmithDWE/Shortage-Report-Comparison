#Copies comments from one excel worksheet to another excel worksheet if the rows are otherwise identical
import tkinter as tk
from tkinter.filedialog import askdirectory
from tkinter import filedialog
from tkinter import messagebox
from tkinter import *
import os
from openpyxl import Workbook
import openpyxl


#File Handling and Setting Sheet 1 as active sheet for both workbooks
root = tk.Tk()
root.withdraw()
print("Select source file using file prompt:")
src_path = filedialog.askopenfilename()
print("Source File: \n"+src_path)
wb_obj_src = openpyxl.load_workbook(src_path)
sheet_obj_src = wb_obj_src["Sheet1"]
print("Select destination file using file prompt:")
dest_path = src_path = filedialog.askopenfilename()
print("Destination File: \n"+dest_path)
wb_obj_dest = openpyxl.load_workbook(dest_path)
sheet_obj_dest = wb_obj_dest["Sheet1"]
#

#Importing columns (import cols C,D,E from both sheets as vectors, match the entries, if an entry matches copy the comment to new sheet)
#To match the text boxes, compares every src to every destination to ensure that none are missed- ineffecient, do not use with large spreadsheets o(n2)
srcC = []
srcD = []
srcE = []
destC = []
destD = []
destE = []
srcT = []

for row in sheet_obj_src:
    val = row[2].value
    srcC.append(val)
for row in sheet_obj_src:
    val = row[3].value
    srcD.append(val)
for row in sheet_obj_src:
    val = row[4].value
    srcE.append(val)
for row in sheet_obj_dest:
    val = row[2].value
    destC.append(val)
for row in sheet_obj_dest:
    val = row[3].value
    destD.append(val)
for row in sheet_obj_dest:
    val = row[4].value
    destE.append(val)
for row in sheet_obj_src:
    val = row[19].value
    srcT.append(val)
    


#create list that stores 1/0 for match or no match for a row's values in C/D/E, then use that vector at the end to compare the two comments cols
match = [0]*sheet_obj_src.max_row
for i in range(len(srcC)):
    for j in range(len(destC)):
        if(srcC[i] == destC[j]):
            match[j] = srcT[i]

for i in range(len(srcD)):
    for j in range(len(destD)):
        if(srcD[i] == destD[j]):
            match[j] = srcT[i]
for i in range(len(srcE)):
    for j in range(len(destE)):
        if(srcE[i] == destE[j]):
            match[j] = srcT[i]


#Writing comments to dest sheet *ALWAYS WRITES TO COl19 (T)
for i in range(len(match)):
    if (match[i] != 0):
        for j in range(1, sheet_obj_dest.max_row+1):
            sheet_obj_dest.cell(row = i+1, column = 20).value = match[i]
sheet_obj_dest.cell(row = 1, column = 20).value = "COPIED COMMENTS"
wb_obj_dest.save(dest_path)
print("Operation Completed.")