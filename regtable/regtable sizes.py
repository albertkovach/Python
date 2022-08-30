from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

from threading import Thread
import time, datetime
import datetime as datetime2

import os, sys, re, csv
from pathlib import Path

import PyPDF2
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter
from PyPDF2 import PdfFileMerger

from pikepdf import Pdf as pike

from pdfminer.layout import LAParams, LTTextBox, LTText, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator
from pdfminer.high_level import extract_text
import fnmatch

from difflib import SequenceMatcher


global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150


global CSVfile
global LOGfile

global CounterWorking
global CounterStartTime




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
    root.geometry('280x160+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()

class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.parent.title("regtable")
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
        
        global OutputNameLbl
        OutputNameLbl = Label(text="Имя файла CSV:", background="white")
        OutputNameLbl.place(x=26, y=62)
        
        global OutputNameEntry
        OutputNameEntry = Entry(fg="black", bg="white", width=20)
        OutputNameEntry.place(x=133, y=60)
        
        global StartBtn
        StartBtn = Button(text='Запуск', command=Run)
        StartBtn.place(x=20, y=55+40, height=20, width=50)
        StartBtn.configure(state = DISABLED)
        
        global ProgressLbl
        ProgressLbl = Label(text="Выберите папку!", background="white")
        ProgressLbl.place(x=20, y=85+40)
        
        global TimeLbl
        TimeLbl = Label(text="", background="white")
        TimeLbl.place(x=90, y=55+40)
        
        
        global CounterWorking
        CounterWorking = False
        
        




def SelectDir():

    global InputDir
    global InputFilesArray
    InputFilesArray = []

    InputDir = ""
    InputDir = filedialog.askdirectory(title='Выберите папку на обработку')
    if InputDir:
        InputDirEntry.configure(state = NORMAL)
        InputDirEntry.delete(0,END)
        InputDirEntry.insert(0,str(Path(InputDir).name))
        InputDirEntry.configure(state = DISABLED)
        
        print('InputDir : {0}'.format(InputDir))
        print('*********************************')
        
        InputFilesArray.clear()
        for path, subdirs, files in os.walk(InputDir):
            for name in files:
                if fnmatch.fnmatch(name, "*.pdf"):
                    InputFilesArray.append(Path(path, name).as_posix())
                    
        StartBtn.configure(state = NORMAL)
        ProgressLbl.config(text = 'Количество файлов в папках: {0}'.format(len(InputFilesArray)))
                
    else:
        print('InputDir not selected')
      
      
def Run():
    BlockGUI(True)

    registrythread = Thread(target=RegistryFillingThread)
    registrythread.start()

    timethread = Thread(target=TimeUpdater)
    timethread.start()




def RegistryFillingThread():

    global CounterWorking
    global CounterStartTime

    CounterWorking = True
    CounterStartTime = time.time()
    
    global CSVfile
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
    #programpath = Path(__file__).resolve().parent
    
    OutputName = OutputNameEntry.get() + ".csv"
    print(OutputName)
    CSVfile = Path(desktop, OutputName)
    
    OutputLogName = OutputNameEntry.get() + "_LOG.txt"
    global LOGfile
    LOGfile = Path(desktop, OutputLogName)
    
    CSVrows = []
    
    try:
        with open(CSVfile, 'w', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(["Папка", "Размер"])
    except:
        messagebox.showerror("", "Невозможно создать отчет CSV !")
        
        
    try:
        with open(LOGfile, 'w', newline='') as log:
            log.writelines('regtable log file' + '\n')
    except:
        messagebox.showerror("", "Невозможно создать log.txt !")
    
    global InputFilesArray
    filenum = 0

    for f in range(len(InputFilesArray)):
        
        fileprecent = int(filenum / len(InputFilesArray) * 100)
        filenum = filenum + 1
        ProgressLbl.config(text = 'Обработка: {0} из {1}. Выполнено: {2}%'.format(filenum, len(InputFilesArray), fileprecent)) 
        
        print("")
        print("")
        print("**********************************************************")
        print('Документ №{0}: {1}'.format(f+1, Path(InputFilesArray[f]).name))
        
        filesize = os.path.getsize(InputFilesArray[f])/1e+6
        CSVrows.append([Path(InputFilesArray[f]).parent.name, filesize])
                                                       
        try:
            with open(CSVfile, 'a', newline='') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerows(CSVrows)
        except:
            line = 'CSV append error. ' + time.strftime('%Y-%m-%d|%H:%M:%S') + ' --- ' + file +'\n'
            with open(LOGfile, 'a', newline='') as log:
                log.writelines(line)
                
        CSVrows = []
        
        
    ProgressLbl.config(text = 'Обработка {0} файлов завершена!'.format(len(InputFilesArray)))
    CounterWorking = False
    BlockGUI(False)


def DocTextSearch(file):

    docnumber = 'нет данных'
    docaddress = 'нет данных'
    
    addressfound = False
    numberfound = False
    
    addressclfound = False
    addressclcounter = 0
    
    global CSVfile  
    CSVrows = []
    
    global LOGfile


    try:
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
                                firstletter = text[0]
                                if firstletter == 'R':
                                    if len(text) == 23 or len(text) == 25:
                                        docnumber = text
                                        numberfound = True
                                        print('============================================')
                                        print('Number: {0}'.format(docnumber))
                                        break
                                        
                    for textbox in layout:
                        if isinstance(textbox, LTText):
                            for line in textbox:
                                text = line.get_text().replace('\n', '')
                                
                                if addressclfound == True:
                                    addressclcounter = addressclcounter + 1
                                    if addressclcounter == 2:
                                        docaddress = text
                                        addressfound = True
                                        print('Address: {0}'.format(docaddress))
                                        print('============================================')
                                        break
                                        
                                if addressclfound == False:
                                    if text[0] == '№' and text[2] == 'п' and text[3] == '/' and text[4] == 'п':
                                        addressclfound = True
    except:
        line = 'File open error. ' + time.strftime('%Y-%m-%d|%H:%M:%S') + ' --- ' + file +'\n'
        with open(LOGfile, 'a', newline='') as log:
            log.writelines(line)

     
    if addressfound == False:
        docaddress = 'Адрес не найден'
    if numberfound == False:
        docaddress = 'Номер не найден'
    CSVrows.append([Path(file).parent.name, docnumber, docaddress])
                                                   
    try:
        with open(CSVfile, 'a', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerows(CSVrows)
    except:
        line = 'CSV append error. ' + time.strftime('%Y-%m-%d|%H:%M:%S') + ' --- ' + file +'\n'
        with open(LOGfile, 'a', newline='') as log:
            log.writelines(line)
        
     


     
def BlockGUI(yes):
    if yes:
        InputDirBtn.configure(state = DISABLED)
        StartBtn.configure(state = DISABLED)
    else:
        InputDirBtn.configure(state = NORMAL)
        StartBtn.configure(state = NORMAL)

def TimeUpdater():
    global CounterWorking
    global CounterStartTime

    while CounterWorking:
        CreateDocTime = time.time()
        result = CreateDocTime - CounterStartTime
        result = datetime2.timedelta(seconds=round(result))
        TimeLbl.config(text = str(result))
        time.sleep(0.01)
        
def ResourcePath(relative_path):    
    try:       
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


if __name__ == '__main__':
    main()
