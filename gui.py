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

    #Save
    wb2.save(str(file2name))

file_btn = Button(master,text="Open",command=opennew).grid(row=1)

#Text (Start & End)
def print_input():
    print(start_t.get() + " " + end_t.get() + " " + date_t.get())

lb_st = Label(master,text="From Where To Start").grid(row=2)
start_t = Entry(master,width=20)
start_t.grid(row=3,sticky=W)

lb_en = Label(master,text="End").grid(row=4)
end_t = Entry(master,width=20)
end_t.grid(row=5,sticky=W)

#Text (Date)
lb_dt = Label(master,text="Enter The Date").grid(row=6)
date_t = Entry(master,width=30)
date_t.grid(row=7,sticky=W)

text_btn = Button(master,text="Print",command=print_input).grid(row=8)

#Customize Msg (Optional)
lb_cm = Label(master,text="Enter The Msg(Optional)").grid(row=9)
cm_t = Entry(master,width=30)
cm_t.grid(row=10,sticky=W)

master.geometry("700x500")
master.mainloop()