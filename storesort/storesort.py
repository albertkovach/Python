from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

import os, sys, re
from pathlib import Path
from threading import Thread
import time
import datetime as datetime2

from pdfminer.layout import LAParams, LTTextBox, LTText, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator


global root


global SorterInputDir
global SorterFilesArray
global SorterInputDirSel
global SorterOutputDirSel
global SorterWorking
global SorterStartTime






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
        global SorterInputDirLbl
        SorterInputDirLbl = Label(text="Выберите файлы на сортировку:", background="white", font=("Arial", 10))
        
        global SorterInputDirEntry
        SorterInputDirEntry = Entry(fg="black", bg="white", width=20)
        SorterInputDirEntry.configure(state = DISABLED)
        
        global SorterInputDirBtn
        SorterInputDirBtn = Button(text='Выбор', command=SorterInputDirChoose)
        
        global SorterInputFilesCountLbl
        SorterInputFilesCountLbl = Label(text="", background="white")
        
        
        global SorterOutputDirLbl
        SorterOutputDirLbl = Label(text="Выберите папку для отсортированных:", background="white", font=("Arial", 10))
        
        global SorterOutputDirEntry
        SorterOutputDirEntry = Entry(fg="black", bg="white", width=20)
        SorterOutputDirEntry.configure(state = DISABLED)
        
        global SorterOutputDirBtn
        SorterOutputDirBtn = Button(text='Выбор', command=SorterOutputDirChoose)
        
        
        global SorterStartBtn
        SorterStartBtn = Button(text='Запуск', command=SorterStart)
        SorterStartBtn.configure(state = DISABLED)
        
        global SorterStatusLbl
        SorterStatusLbl = Label(text="", background="white")
        
        global SorterTimeLbl
        SorterTimeLbl = Label(text="", background="white")


        SorterInputDirLbl.place        (x=16, y=7+40)
        SorterInputDirEntry.place      (x=20, y=30+40, width=275)
        SorterInputDirBtn.place        (x=305, y=29+40, height=20)
        SorterInputFilesCountLbl.place (x=16, y=49+40)
        SorterOutputDirLbl.place       (x=17, y=87+35)
        SorterOutputDirEntry.place     (x=75, y=111+35, width=275)
        SorterOutputDirBtn.place       (x=20, y=110+35, height=20)
        SorterStartBtn.place           (x=20, y=190, width=85)
        SorterStatusLbl.place          (x=120, y=183)
        SorterTimeLbl.place            (x=120, y=200)
        
        
        global SorterInputDir
        global SorterFilesArray
        global SorterInputDirSel
        global SorterOutputDirSel
        global SorterWorking
        global SorterStartTime
        
        SorterInputDir = ""
        SorterFilesArray = []
        
        SorterInputDirSel = False
        SorterOutputDirSel = False
        
        SorterWorking = False
        SorterStartTime = ""
        



def SorterInputDirChoose():
    global SorterInputDir
    global SorterFilesArray
    global SorterInputDirSel
    global SorterOutputDirSel

    SorterInputDir = filedialog.askdirectory(title="Выберите папку на сортировку")
    
    if SorterInputDir:
        SorterInputDirEntry.configure(state = NORMAL)
        SorterInputDirEntry.delete(0,END)
        SorterInputDirEntry.insert(0,str(Path(SorterInputDir).name))
        SorterInputDirEntry.configure(state = DISABLED)
        
        print('SorterInputDir:', SorterInputDir)
        
        SorterFilesArray.clear()
        for file in os.listdir(SorterInputDir):
            if file.endswith(".pdf"):
                SorterFilesArray.append(os.path.join(SorterInputDir, file))
        
        if len(SorterFilesArray) == 0:
            SorterInputFilesCountLbl.config(text = 'Нет файлов PDF !')
            print('no pdf in folder !')
            SorterInputDirSel = False
        else:
            SorterInputFilesCountLbl.config(text = 'Количество файлов PDF: ' + str(len(SorterFilesArray)))
            print('number of valid files -', str(len(SorterFilesArray)))
            SorterInputDirSel = True
    else:
        print('SorterInputDir is NOT defined')
        
        
    if SorterInputDirSel and SorterOutputDirSel:
        SorterStartBtn.configure(state = NORMAL)
    else:
        SorterStartBtn.configure(state = DISABLED)




def SorterOutputDirChoose():
    global SorterOutputDir
    global SorterInputDirSel
    global SorterOutputDirSel
    
    SorterOutputDir = filedialog.askdirectory(title="Выберите папку для отсортированных")
    if SorterInputDir:
        SorterOutputDirEntry.configure(state = NORMAL)
        SorterOutputDirEntry.delete(0,END)
        SorterOutputDirEntry.insert(0,str(Path(SorterOutputDir).name))
        SorterOutputDirEntry.configure(state = DISABLED)
        
        SorterOutputDirSel = True
        print('SorterOutputDir:', SorterOutputDir)
    else:
        print('SorterInputDir is NOT defined')
        
        
    if SorterInputDirSel and SorterOutputDirSel:
        SorterStartBtn.configure(state = NORMAL)
    else:
        SorterStartBtn.configure(state = DISABLED)




def SorterStart():

    SorterMainThreadthread = Thread(target=SorterMainThread)
    SorterMainThreadthread.start()
    timethread = Thread(target=TimeUpdater)
    timethread.start()




def SorterMainThread():
    global SorterWorking
    global SorterStartTime

    global SorterInputDir
    global SorterOutputDir
    global SorterFilesArray
    
    SorterStartTime = time.time()
    SorterWorking = True
    SorterBlockGUI(True)
    
    for f in range(len(SorterFilesArray)):
        print("_________________________")
        print('Документ №{0}: {1}'.format(f+1, Path(SorterFilesArray[f]).name))
        
        invoicedata = SorterFileTextSearch(SorterFilesArray[f])
        if isinstance(invoicedata, list):
            statustext = "Документ {0} из {1}".format(f, len(SorterFilesArray))
            SorterStatusLbl.config(text = str(statustext))
            print('ИНН, КПП документа: {0}'.format(invoicedata))
            
            fileoutputdir = Path(SorterOutputDir, invoicedata[1])
            
            
            
        else:
            print('Не найдено ИНН, КПП !')
            
            
            
            
    #isSorterInputDir = os.path.isdir(SorterInputDir)
    
    
    
    
    
    SorterStatusLbl.config(text = "Обработка завершена !")
    SorterWorking = False
    SorterBlockGUI(False)




def SorterFileTextSearch(file):

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




def SorterBlockGUI(yes):
    if yes:
        SorterInputDirBtn.configure(state = DISABLED)
        SorterOutputDirBtn.configure(state = DISABLED)
        SorterStartBtn.configure(state = DISABLED)
    else:
        SorterInputDirBtn.configure(state = NORMAL)
        SorterOutputDirBtn.configure(state = NORMAL)
        SorterStartBtn.configure(state = NORMAL)




def TimeUpdater():
    global SorterWorking
    global SorterStartTime

    time.sleep(0.01)
    while SorterWorking:
        CreateDocTime = time.time()
        result = CreateDocTime - SorterStartTime
        result = datetime2.timedelta(seconds=round(result))
        SorterTimeLbl.config(text = str(result))
        time.sleep(0.01)


if __name__ == '__main__':
    main()
