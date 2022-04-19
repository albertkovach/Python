from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

global root
global scrnwparam
global scrnhparam
scrnwparam = 140
scrnhparam = 100


import PyPDF2
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter
from PyPDF2 import PdfFileMerger

import csv

from pikepdf import Pdf as pike

from pathlib import Path
import re
import os
import sys
import time
import datetime as datetime2
from threading import Thread

from pdfminer.high_level import extract_text
import fnmatch

global InputDir
global InputFilesArray
global IsInputDirSel
global CounterWorking
global CounterStartTime
global SubdirMode
global CountWordsMode
global AllPageCount
global AllWordsCount
global CSVfile





def main():
    global root

    root = Tk()
    root.resizable(False, False)
    
    datafile = "icon.ico"
    if not hasattr(sys, "frozen"):
        datafile = os.path.join(os.path.dirname(__file__), datafile)
    else:
        datafile = os.path.join(sys.prefix, datafile)
    root.iconbitmap(default=ResourcePath(datafile))
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('280x130+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()


class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.parent.title("PageCounter")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        global CountWordsMode
        CountWordsMode = IntVar()
        CountWordsMode.set(0)
        
        global SubdirMode
        SubdirMode = IntVar()
        SubdirMode.set(1)
        
        global IsInputDirSel
        IsInputDirSel = False
        
        global CounterWorking
        CounterWorking = False
        
        global InputFilesArray
        InputFilesArray = []
        
        
        
        global MainLbl
        MainLbl = Label(text="Выберите папку:", background="white")
        MainLbl.place(x=16, y=10)
        
        global InputDirEntry
        InputDirEntry = Entry(fg="black", bg="white", width=30)
        InputDirEntry.place(x=20, y=32)
        InputDirEntry.configure(state = DISABLED)
        
        global FilesCountLbl
        FilesCountLbl = Label(text="", background="white")
        FilesCountLbl.place(x=16, y=52)
        
        global SelectFileBtn
        SelectFileBtn = Button(text='Выбор', command=SelectDir)
        SelectFileBtn.place(x=210, y=30, height=20)
        
        global ModeSubdircbtn
        ModeSubdircbtn = Checkbutton(text="По вложенным", background="white", variable=SubdirMode, onvalue=1, offvalue=0, command=FilesCount)
        ModeSubdircbtn.place(x=150, y=63, height=20)
        
        global ModeCountWordscbtn
        ModeCountWordscbtn = Checkbutton(text="Cчитать слова", background="white", variable=CountWordsMode, onvalue=1, offvalue=0)
        ModeCountWordscbtn.place(x=150, y=85, height=20)
        
        global OpenReportBtn
        OpenReportBtn = Button(text='Открыть отчет', command=OpenReport)
        
        global CountPagesBtn
        CountPagesBtn = Button(text='Посчитать', command=CounterStart)
        CountPagesBtn.place(x=20, y=83, width=100, height=25)
        
        
        
        
        global ProgressMainLbl
        ProgressMainLbl = Label(text="", background="white")
        ProgressMainLbl.place(x=16, y=115)
        
        global ProgressFileLbl
        ProgressFileLbl = Label(text="", background="white")
        ProgressFileLbl.place(x=16, y=135)
        
        global ProgressPagesLbl
        ProgressPagesLbl = Label(text="", background="white")
        ProgressPagesLbl.place(x=16, y=155)
        
        global TimeLbl
        TimeLbl = Label(text="", background="white")
        TimeLbl.place(x=210, y=155)



def SelectDir():
    print('===================')
    print('==SelectDir started')
    
    global InputDir
    global InputFilesArray
    global IsInputDirSel

    InputDir = ""
    InputDir = filedialog.askdirectory(title='Выберите папку на обработку')
    if InputDir:
        InputDirEntry.configure(state = NORMAL)
        InputDirEntry.delete(0,END)
        InputDirEntry.insert(0,str(InputDir))
        InputDirEntry.configure(state = DISABLED)
        print('IDC: InputDir : {0}'.format(InputDir))
        
        FilesCountLbl.config(text = ('Поиск фалов PDF...'))
        filescanthread = Thread(target=FilesCount)
        filescanthread.start()

        ProgressMainLbl.config(text = "")  
        ProgressFileLbl.config(text = "")  
        ProgressPagesLbl.config(text = "")
        TimeLbl.config(text = "")
        
        OpenReportBtn.place_forget()
        ResizeGUI(False)
    else:
        print('IDC: InputFile not selected')

    print('==SelectDir ended')
    print('===================')


def FilesCount():
    global SubdirMode
    global InputDir
    global InputFilesArray
    global IsInputDirSel
    
    InputFilesArray.clear()
    print('===================')
    print('==FilesCount started')
    if SubdirMode.get() == 1:
        for path, subdirs, files in os.walk(InputDir):
            for name in files:
                if fnmatch.fnmatch(name, "*.pdf"):
                    InputFilesArray.append(Path(path, name).as_posix())
    else:
        for f in os.listdir(InputDir):
            if f.endswith(".pdf"):
                InputFilesArray.append(os.path.join(InputDir, f))
                
                
    print('IDC: files count: {0}'.format(len(InputFilesArray)))
    
    if len(InputFilesArray) > 0:
        FilesCountLbl.config(text = ('Файлов PDF: ' + str(len(InputFilesArray))))
        print('IDC: files list:')
        for i in range(len(InputFilesArray)):
            print('IDC: f {0}: {1}'.format(i, InputFilesArray[i]))
        
        IsInputDirSel = True
        
    else:
        FilesCountLbl.config(text = ('Нет фалов PDF !'))
        print('IDC: no pdf files !')
        IsInputDirSel = False
        
    print('==FilesCount ended')
    print('===================')


def CounterStart():
    global IsInputDirSel
    OpenReportBtn.place_forget()

    if IsInputDirSel:
        ResizeGUI(True)
    
        counterthread = Thread(target=Counter)
        counterthread.start()
        timethread = Thread(target=TimeUpdater)
        timethread.start()
    else:
        FilesCountLbl.config(text = ('Выбери папку !'))



def Counter():
    global InputDir
    global CounterWorking
    global CounterStartTime
    global InputFilesArray
    global CountWordsMode
    global SubdirMode
    global AllPageCount
    global AllWordsCount
    global CSVfile
    
    
    CounterWorking = True
    CounterStartTime = time.time()
    BlockGUI(True)
    
    
    programpath = Path(__file__).resolve().parent
    CSVfile = Path(programpath, "counter.csv")
    CSVrows = []

    AllPageCount = 0
    AllWordsCount = 0
    CurrentPageCount = 0
    CurrentWordsCount = 0
    
    
    for f in range (len(InputFilesArray)):
        ProgressMainLbl.config(text = "Обработка файла {0} из {1}".format(f+1, len(InputFilesArray)))
        ProgressFileLbl.config(text = "Подсчет страниц...")
        #CurrentPageCount = CountPages(InputFilesArray[f])
        CurrentPageCount = PikeCountPages(InputFilesArray[f])
        AllPageCount = AllPageCount + CurrentPageCount
        
        if CountWordsMode.get() == 1:
            ProgressFileLbl.config(text = "Подсчет слов...")
            CurrentWordsCount = CountWords(InputFilesArray[f])
            AllWordsCount = AllWordsCount + CurrentWordsCount
            ProgressPagesLbl.config(text = "Всего стр: {0}, слов: {1}".format(AllPageCount, AllWordsCount))
        else:
            ProgressPagesLbl.config(text = "Всего стр: {0}".format(AllPageCount))
            
        if CountWordsMode.get() == 1:
            if SubdirMode.get() == 1:
                CSVrows.append([InputFilesArray[f], CurrentPageCount, CurrentWordsCount])
            else:
                CSVrows.append([Path(InputFilesArray[f]).name, CurrentPageCount, CurrentWordsCount])
        else:
            if SubdirMode.get() == 1:
                CSVrows.append([InputFilesArray[f], CurrentPageCount])
            else:
                CSVrows.append([Path(InputFilesArray[f]).name, CurrentPageCount])
            
    try:
        with open(CSVfile, 'w', newline='') as f:
            writer = csv.writer(f, delimiter=';')

            writer.writerow(["Папка:"])
            writer.writerow([str(InputDir)])
            writer.writerow(["", "Файлов :", str(len(InputFilesArray))])
            writer.writerow(["", "Страниц :", str(AllPageCount)])
            
            if CountWordsMode.get() == 1:
                writer.writerow(["", "Слов:", str(AllWordsCount)])
                writer.writerow([""])
                writer.writerow(["Файл", "Страниц", "Слов"])
            else:
                writer.writerow([""])
                writer.writerow(["Файл", "Страниц"])
                
            writer.writerows(CSVrows)
    except:
        messagebox.showerror("", "Невозможно создать отчет CSV !")

    OpenReportBtn.place(x=158, y=132, height=20)

    ProgressMainLbl.config(text = "")  
    ProgressFileLbl.config(text = "Завершено !")
    
    CounterWorking = False
    BlockGUI(False)
    
    messagebox.showinfo("", "Обработка завершена !")


def OpenReport():
    global CSVfile
    
    fileexists = os.path.isfile(CSVfile)
    if fileexists:
        os.startfile(CSVfile)


def BlockGUI(yes):
    if yes:
        CountPagesBtn.configure(state = DISABLED)
        SelectFileBtn.configure(state = DISABLED)
        ModeCountWordscbtn.configure(state = DISABLED)
        ModeSubdircbtn.configure(state = DISABLED)
    else:
        CountPagesBtn.configure(state = NORMAL)
        SelectFileBtn.configure(state = NORMAL)
        ModeCountWordscbtn.configure(state = NORMAL)
        ModeSubdircbtn.configure(state = NORMAL)


def CountWords(file): 
    text = extract_text(Path(file).as_posix())
    words = re.findall(r'\b\S+\b', text)
    #for i in range (len(words)):
    #    print(words[i])
    #print(len(words))
    return len(words)
    
def PikeCountPages(file):
    pdf_doc = pike.open(file)
    pdf_page_count = len(pdf_doc.pages)
    return pdf_page_count
    
def CountPages(file):
    pdf = PdfFileReader(file)
    try:
        pagecount = pdf.getNumPages()
    except:
        msgbxlbl = ['Ошибка открытия файла !', str(file)]
        messagebox.showerror("", "\n".join(msgbxlbl))
        return 0
    pdf = ""
    return pagecount
    
    
def TimeUpdater():
    global CounterWorking
    global CounterStartTime

    while CounterWorking:
        CreateDocTime = time.time()
        result = CreateDocTime - CounterStartTime
        result = datetime2.timedelta(seconds=round(result))
        TimeLbl.config(text = str(result))
        time.sleep(0.01)
    

def ResizeGUI(full):
    global scrnwparam
    global scrnhparam
    
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    
    if full:
        root.geometry('280x190+{}+{}'.format(scrnw, scrnh))
    else:
        root.geometry('280x130+{}+{}'.format(scrnw, scrnh))


def ResourcePath(relative_path):    
    try:       
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)



if __name__ == '__main__':
    main()
