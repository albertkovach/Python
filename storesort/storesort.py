from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

import os, sys, re, shutil
from pathlib import Path
from threading import Thread
import time
import datetime as datetime2

from pdfminer.layout import LAParams, LTTextBox, LTText, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator


global root


global PlaceInputDir
global PlaceFilesArray
global PlaceInputDirSel
global PlaceOutputDirSel
global PlaceWorking
global PlaceStartTime






def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('375x250+{}+{}'.format(scrnw, scrnh))
        
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
        global PlaceInputDirLbl
        PlaceInputDirLbl = Label(text="Выберите файлы на сортировку:", background="white", font=("Arial", 10))
        
        global PlaceInputDirEntry
        PlaceInputDirEntry = Entry(fg="black", bg="white", width=20)
        PlaceInputDirEntry.configure(state = DISABLED)
        
        global PlaceInputDirBtn
        PlaceInputDirBtn = Button(text='Выбор', command=PlaceInputDirChoose)
        
        global PlaceInputFilesCountLbl
        PlaceInputFilesCountLbl = Label(text="", background="white")
        
        
        global PlaceOutputDirLbl
        PlaceOutputDirLbl = Label(text="Выберите папку для отсортированных:", background="white", font=("Arial", 10))
        
        global PlaceOutputDirEntry
        PlaceOutputDirEntry = Entry(fg="black", bg="white", width=20)
        PlaceOutputDirEntry.configure(state = DISABLED)
        
        global PlaceOutputDirBtn
        PlaceOutputDirBtn = Button(text='Выбор', command=PlaceOutputDirChoose)
        
        
        global PlaceStartBtn
        PlaceStartBtn = Button(text='Запуск', command=PlaceStart)
        PlaceStartBtn.configure(state = DISABLED)
        
        global PlaceStatusLbl
        PlaceStatusLbl = Label(text="", background="white")
        
        global PlaceTimeLbl
        PlaceTimeLbl = Label(text="", background="white")


        PlaceInputDirLbl.place        (x=16, y=7+40)
        PlaceInputDirEntry.place      (x=20, y=30+40, width=275)
        PlaceInputDirBtn.place        (x=305, y=29+40, height=20)
        PlaceInputFilesCountLbl.place (x=16, y=49+40)
        PlaceOutputDirLbl.place       (x=17, y=87+35)
        PlaceOutputDirEntry.place     (x=75, y=111+35, width=275)
        PlaceOutputDirBtn.place       (x=20, y=110+35, height=20)
        PlaceStartBtn.place           (x=20, y=190, width=85)
        PlaceStatusLbl.place          (x=120, y=183)
        PlaceTimeLbl.place            (x=120, y=200)
        
        
        global PlaceInputDir
        global PlaceFilesArray
        global PlaceInputDirSel
        global PlaceOutputDirSel
        global PlaceWorking
        global PlaceStartTime
        
        PlaceInputDir = ""
        PlaceFilesArray = []
        
        PlaceInputDirSel = False
        PlaceOutputDirSel = False
        
        PlaceWorking = False
        PlaceStartTime = ""





def PlaceInputDirChoose():
    global PlaceInputDir

    PlaceInputDir = filedialog.askdirectory(title="Выберите папку на сортировку")
    
    if PlaceInputDir:
        PlaceInputDirEntry.configure(state = NORMAL)
        PlaceInputDirEntry.delete(0,END)
        PlaceInputDirEntry.insert(0,str(Path(PlaceInputDir).name))
        PlaceInputDirEntry.configure(state = DISABLED)
        
        print('PlaceInputDir:', PlaceInputDir)
        
        PlaceInputDirCheck()
    else:
        print('PlaceInputDir is NOT defined')
        

def PlaceInputDirCheck():
    global PlaceInputDir
    global PlaceFilesArray
    global PlaceInputDirSel
    global PlaceOutputDirSel

    PlaceFilesArray.clear()
    for file in os.listdir(PlaceInputDir):
        if file.endswith(".pdf"):
            PlaceFilesArray.append(os.path.join(PlaceInputDir, file))
    
    if len(PlaceFilesArray) == 0:
        PlaceInputFilesCountLbl.config(text = 'Нет файлов PDF !')
        print('no pdf in folder !')
        PlaceInputDirSel = False
    else:
        PlaceInputFilesCountLbl.config(text = 'Количество файлов PDF: ' + str(len(PlaceFilesArray)))
        print('number of valid files -', str(len(PlaceFilesArray)))
        PlaceInputDirSel = True

    if PlaceInputDirSel and PlaceOutputDirSel:
        PlaceStartBtn.configure(state = NORMAL)
    else:
        PlaceStartBtn.configure(state = DISABLED)


def PlaceOutputDirChoose():
    global PlaceOutputDir
    global PlaceInputDirSel
    global PlaceOutputDirSel
    
    PlaceOutputDir = filedialog.askdirectory(title="Выберите папку для отсортированных")
    if PlaceInputDir:
        PlaceOutputDirEntry.configure(state = NORMAL)
        PlaceOutputDirEntry.delete(0,END)
        PlaceOutputDirEntry.insert(0,str(Path(PlaceOutputDir).name))
        PlaceOutputDirEntry.configure(state = DISABLED)
        
        PlaceOutputDirSel = True
        print('PlaceOutputDir:', PlaceOutputDir)
    else:
        print('PlaceInputDir is NOT defined')
        
        
    if PlaceInputDirSel and PlaceOutputDirSel:
        PlaceStartBtn.configure(state = NORMAL)
    else:
        PlaceStartBtn.configure(state = DISABLED)


def PlaceStart():

    PlaceMainThreadthread = Thread(target=PlaceMainThread)
    PlaceMainThreadthread.start()
    timethread = Thread(target=PlaceTimeUpdater)
    timethread.start()


def PlaceMainThread():
    global PlaceWorking
    global PlaceStartTime

    global PlaceInputDir
    global PlaceOutputDir
    global PlaceFilesArray
    
    PlaceStartTime = time.time()
    PlaceWorking = True
    PlaceBlockGUI(True)
    
    for f in range(len(PlaceFilesArray)):
        print("_________________________")
        print('Документ №{0}: {1}'.format(f+1, Path(PlaceFilesArray[f]).name))
        
        invoicedata = PlaceFileTextSearch(PlaceFilesArray[f])
        if isinstance(invoicedata, list):
            statustext = "Обработка {0} из {1}".format(f, len(PlaceFilesArray))
            PlaceStatusLbl.config(text = str(statustext))
            print('ИНН, КПП документа: {0}'.format(invoicedata))
            
            fileoutputdir = Path(PlaceOutputDir, invoicedata[1])
            fileoutputdirexist = os.path.exists(fileoutputdir)
            print('fileoutputdir: {0}, exists: {1}'.format(fileoutputdir, fileoutputdirexist))
            
            if not fileoutputdirexist:
                os.makedirs(fileoutputdir)
                
            fileoutputpath = Path(fileoutputdir, Path(PlaceFilesArray[f]).name).as_posix()
            shutil.move(PlaceFilesArray[f], fileoutputpath)
            
            
        else:
            print('Не найдено ИНН, КПП !')
            msgbxlbl = ['В документе не найдено ИНН, КПП !', '{0}'.format(PlaceFilesArray[f])]
            messagebox.showerror("", "\n".join(msgbxlbl))
            
    
    PlaceStatusLbl.config(text = "Обработка завершена !")
    PlaceWorking = False
    PlaceBlockGUI(False)
    PlaceInputDirCheck()


def PlaceFileTextSearch(file):

    with open(file, 'rb') as pdftomine:
        manager = PDFResourceManager()
        laparams = LAParams()
        dev = PDFPageAggregator(manager, laparams=laparams)
        interpreter = PDFPageInterpreter(manager, dev)
        pages = PDFPage.get_pages(pdftomine)

        for pagenumber, page in enumerate(pages):
            if pagenumber == 0:
                interpreter.process_page(page)
                layout = dev.get_result()
                
                for textbox in layout:
                    if isinstance(textbox, LTText):
                        for line in textbox:
                            text = line.get_text().replace('\n', '')
                            if len(text) == 22 or len(text) == 20:
                                invoiceinn = re.sub("[^0-9]", "", (text.partition("/")[0]))
                                invoicekpp = re.sub("[^0-9]", "", (text.partition("/")[2]))
                                if invoiceinn.isnumeric() and invoicekpp.isnumeric():
                                    if len(invoiceinn)==10 and len(invoicekpp)==9:
                                        #print("_________________________")
                                        #print('ИНН документа: {0}'.format(invoiceinn))
                                        #print('КПП документа: {0}'.format(invoicekpp))
                                        invoicedata = [invoiceinn, invoicekpp]
                                        return invoicedata
                                        break
        return "NONE"


def PlaceBlockGUI(yes):
    if yes:
        PlaceInputDirBtn.configure(state = DISABLED)
        PlaceOutputDirBtn.configure(state = DISABLED)
        PlaceStartBtn.configure(state = DISABLED)
    else:
        PlaceInputDirBtn.configure(state = NORMAL)
        PlaceOutputDirBtn.configure(state = NORMAL)
        PlaceStartBtn.configure(state = NORMAL)


def PlaceTimeUpdater():
    global PlaceWorking
    global PlaceStartTime

    time.sleep(0.01)
    while PlaceWorking:
        CreateDocTime = time.time()
        result = CreateDocTime - PlaceStartTime
        result = datetime2.timedelta(seconds=round(result))
        PlaceTimeLbl.config(text = str(result))
        time.sleep(0.01)




if __name__ == '__main__':
    main()
