from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

import os, sys, re, shutil, psutil
from pathlib import Path
from PyPDF2 import PdfFileMerger, PdfFileReader

from threading import Thread

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150


global InputDir
global SubdirArray

def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('280x120+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()

class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.parent.title("")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        global InputDirLbl
        InputDirLbl = Label(text="Директория с папками:", background="white")
        InputDirLbl.place(x=16, y=10)
        
        global InputDirEntry
        InputDirEntry = Entry(fg="black", bg="white", width=30)
        InputDirEntry.place(x=20, y=32)
        
        global InputDirBtn
        InputDirBtn = Button(text='Выбор', command=SelectDir)
        InputDirBtn.place(x=210, y=31, height=20)
        
        global StartBtn
        StartBtn = Button(text='Запуск', command=Run)
        StartBtn.place(x=20, y=55, height=20, width=50)
        StartBtn.configure(state = DISABLED)
        
        global ProgressLbl
        ProgressLbl = Label(text="Выберите директорию!", background="white")
        ProgressLbl.place(x=20, y=85)
        
        global TimeLbl
        TimeLbl = Label(text="", background="white")
        TimeLbl.place(x=90, y=55)
        
        
        global CounterWorking
        CounterWorking = False




def SelectDir():
    global InputDir
    global SubdirArray
    SubdirArray = []

    InputDir = ""
    InputDir = filedialog.askdirectory(title='Выберите папку на обработку')
    if InputDir:
        InputDirEntry.configure(state = NORMAL)
        InputDirEntry.delete(0,END)
        InputDirEntry.insert(0,str(InputDir))
        InputDirEntry.configure(state = DISABLED)
        print('InputDir : {0}'.format(InputDir))
        
        SubdirArray.clear()
        for subdir in os.listdir(InputDir):
            d = Path(InputDir, subdir)
            if os.path.isdir(d):
                SubdirArray.append(d)
                print('**** subdir: {0}'.format(d))
                
        StartBtn.configure(state = NORMAL)
        ProgressLbl.config(text = 'Количество папок: {0}'.format(len(SubdirArray)))
    else:
        print('InputDir not selected')




def Run():
    BlockGUI(True)

    mainthread = Thread(target=Thr)
    mainthread.start()



def Thr():
    global InputDir
    global SubdirArray
    
    print('======= Start =======')
    for i in range(0, len(SubdirArray)):
        MergeAndMove(SubdirArray[i])
        
        
        

def MergeAndMove(DIR):
    print('DIR: {0}'.format(DIR))
    pdfmerger = PdfFileMerger()
    
    FilesArray = []
    
    for file in os.listdir(DIR):
        if file.endswith(".pdf"):
            filename = os.path.join(DIR, file)
            FilesArray.append(filename)
            print('******* file: {0}'.format(filename))
            
    pdfmerger = PdfFileMerger()
    for i in range(0, len(FilesArray)):
        #pdfmerger.append(PdfFileReader(open(FilesArray[i], 'rb')))
        pdfmerger.append(FilesArray[i])
        pdfmerger.append(FilesArray[i])
        
    outputfile = Path(DIR, "на печать.pdf")
    revisedoutputfile = outputfile.as_posix()
    print('Writing file: {0}'.format(filename))
    pdfmerger.write(revisedoutputfile)
    
    pdfmerger = ""
    print('Writing COMPLETE !')


    BlockGUI(False)
    
    
def BlockGUI(yes):
    if yes:
        InputDirBtn.configure(state = DISABLED)
        StartBtn.configure(state = DISABLED)
    else:
        InputDirBtn.configure(state = NORMAL)
        StartBtn.configure(state = NORMAL)
        

if __name__ == '__main__':
    main()
