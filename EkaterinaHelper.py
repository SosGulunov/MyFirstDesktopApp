import main as Eh
import tkinter as tk
from tkinter import END, ttk
import os
import xlrd

window = tk.Tk()
window.title("TimeTable 0.1")
window.geometry('715x440')
window.resizable(width=False, height=False)
check = False
frame_result = tk.Frame(window, width = 358, height=440, bg='white', bd = 1, relief = "solid")
frame_worklist = tk.Frame(window, width = 90, height=440, bg='thistle1')
frame_worklistEntry = tk.Frame(window, width = 91, height=440, bg='thistle1')
frame_restlist = tk.Frame(window, width = 176, height=440, bg='light cyan', bd = 1, relief = "solid")

frame_result.grid(row = 0, column=0)
frame_worklist.grid(row = 0, column=1)
frame_worklistEntry.grid(row = 0, column=2)
frame_restlist.grid(row = 0, column=3)



Txt = [tk.Entry(frame_worklistEntry, width=3, relief = "flat") for i in range(len(Eh.Cheliki))]
Txt2 = [tk.Entry(frame_worklistEntry, width=3, relief = "flat") for i in range(len(Eh.Cheliki))]
lblStd = [tk.Label(frame_worklist, relief = "flat", bg="thistle1", text=Eh.Cheliki[i].name)for i in range(len(Eh.Cheliki))]

def openf():
    os.startfile(r'"Result.xlsx"')
def opentxt():
    os.startfile(r'"Workers.txt"')
    window.quit()

def save_time():
    for i in range(len(Eh.Cheliki)):
        Eh.Cheliki[i].a = int(Txt[i].get())
        Eh.Cheliki[i].b = int(Txt2[i].get())
def createTable():
    global check
    if check == False:
        save_time()
        check = True
    Eh.main()
    #for i in range(len(Eh.Cheliki)):
    #    print(Eh.Cheliki[i].name, Eh.Cheliki[i].a, Eh.Cheliki[i].b)
    table = ttk.Treeview(frame_result, show='headings')
    heads = ['Time','Инфа','Матем','Общая']
    table['columns'] = heads
    excel_workbook = xlrd.open_workbook(r'Result.xlsx')
    excel_sheet = excel_workbook.sheet_by_index(0)
    for header in heads:
        table.heading(header, text=header, anchor='center')
        table.column(header, anchor='center', width=80)
    topList = []
    for j in range (4):
        ls = ''
        ls = [str(excel_sheet.cell_value(i,j)) for i in range(excel_sheet.nrows) if not 'module' in str(excel_sheet.cell_value(i,j))]
        ls.pop(0)
        ls.reverse()
        topList.append(ls)
    lastTopList= [[topList[j][i] for j in range(len(topList))] for i in range(len(topList[0])-1,-1,-1)]
    for row in lastTopList:
        table.insert('',tk.END, values = row)
    table.place(x=20,y=10)

def goRest():
    Name = txtResting.get()
    for i in range(len(Eh.Cheliki)):
        if Eh.Cheliki[i].name==Name:
            Eh.RestCheliki.append(Eh.Cheliki[i])
            Eh.Cheliki.pop(i)
        if i == len(Eh.Cheliki)-1:
            break
    for i in range(len(Eh.RestCheliki)):
        Restlbl = tk.Label(frame_restlist,relief = "flat", bg="light cyan", text=Eh.RestCheliki[i].name)
        Restlbl.place(x = 10, y = 50 + i*25)
    ReWork()
def ReWork():
    for widget in frame_worklist.winfo_children():
        widget.destroy()
    for widget in frame_worklistEntry.winfo_children():
        widget.place(x = 0, y = 500)
        widget.insert(0,'0')
    lbl = [tk.Label(frame_worklist,relief = "flat", bg="thistle1", text=Eh.Cheliki[i].name)for i in range(len(Eh.Cheliki))]
    CreateWorkList(lbl)
def clear_RestList():
    global check
    check = False
    for widget in frame_restlist.winfo_children():
        widget.destroy()
    Eh.RestCheliki.clear()
    Eh.Cheliki = [Eh.Chelik(Eh.n[i]) for i in range((len(Eh.n)))]
    ReWork()
def CreateWorkList(TmpLbl):
    Lbl = TmpLbl
    for i in range(len(Eh.Cheliki)):
        Txt[i].place(x = 20, y = 27 + i*23)
        Txt[i].delete("0", END)
        Txt2[i].place(x = 60, y = 27 + i*23)
        Txt2[i].delete("0", END)

        Lbl[i].place(x = 4, y = 25 + i*23)

list_working = tk.Label(window, width=24, height=1, text="Отдыхают", bg="light cyan" )
list_working.place(x=539, y=1)
txtResting = tk.Entry(window, width=20,)
txtResting.place(x = 555, y = 27)

bAddinRestList = tk.Button(
    text="+",
    width=2,
    height=0,
    bg="white",
    fg="black", bd=1, relief='solid',
    command=goRest,
)
bAddinRestList.place(x = 685, y = 25)
bCreateTimeTable = tk.Button(
    text="Создать расписание",
    width=25,
    height=1,
    bg="white",
    fg="black",
    bd=1,
    relief = "solid",
    command = createTable,
)
bCreateTimeTable.place(x = 357, y = 416)
bClearRestList = tk.Button(
    text="Никто больше не отдыхает!",
    width=24,
    height=1,
    bg="white",
    fg="black",
    bd=1,
    relief = "solid",
    command=clear_RestList,
)
bClearRestList.place(x = 539, y = 416)
bOpen = tk.Button(
    text="Открыть",
    width=50,
    height=1,
    bg="light cyan",
    fg="black",
    bd=1,
    relief="solid",
    command=openf,
)
bOpen.place(x = 0, y = 416)
bOpenTxt = tk.Button(
    window, width=25, height=0, text="Работают", bg="thistle1", bd=1, relief='solid',
    command = opentxt,
)
bOpenTxt.place(x=357, y=0)
CreateWorkList(lblStd)
window.mainloop()