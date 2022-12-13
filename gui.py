from tkinter import *
import openpyxl as xl;
from tkinter import filedialog

master=Tk()

#File
def opennew():
    #copy
    master.filename = filedialog.askopenfilename()
    wb1 = xl.load_workbook(master.filename) 
    ws1 = wb1.worksheets[0]

    #paste
    file2name = "finalsheet.xlsx"
    wb2 = xl.load_workbook(file2name) 
    ws2 = wb2.active
    mr=ws1.max_row
    mc=ws1.max_column

    for i in range (1,mr+1):
        for j in range (1,mc+1):
            c=ws1.cell(row=i,column=j)

            ws2.cell(row=i,column=j).value=c.value

    wb2.save(str(file2name))
my_btn = Button(master,text="Open",command=opennew).grid(row=3)

#Text


master.geometry("700x500")
master.mainloop()