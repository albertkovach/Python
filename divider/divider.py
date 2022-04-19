from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
import ctypes as ct

from threading import Thread

import time
import datetime as datetime2
from datetime import datetime

from pathlib import Path
import shutil
import os
import sys

import PyPDF2
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter
from PyPDF2 import PdfFileMerger

from pdfminer.layout import LAParams, LTTextBox, LTText, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator

from difflib import SequenceMatcher




# >>>>> GUI variables <<<<<
# -------------------------
global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

global CurrentUIpage
# -------------------------



# >>> Divide variables <<<
# -------------------------
global DivideInputFilesArray
global DivideInputFile
global DivideIsInputSel
global DivideReadyToGo

global DivideOutputDir
global DivideIsOutputSel

global DivideAllPagesCount
global DivideFirstPagesArray

global DivideIsRunning
global DivideStartedTime
# -------------------------



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
    root.geometry('375x310+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()



class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.parent.title("Divider")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        # >> Divide var init <<
        # ----------------------
        global DivideInputFilesArray
        global DivideInputFile
        global DivideOutputDir
        DivideInputFilesArray = []
        DivideInputFile = ""  
        DivideOutputDir = ""  
                
        global DivideIsInputSel
        global DivideIsOutputSel
        global DivideIsRunning
        global DivideReadyToGo
        DivideIsInputSel = False
        DivideIsOutputSel = False
        DivideIsRunning = False
        DivideReadyToGo = False
        
        global DivideFirstPagesArray
        DivideFirstPagesArray = []
        # ----------------------
    
    
    
        # >>>>>>>>>>>>>>> Mode header <<<<<<<<<<<<<<<
        # ===========================================
        global MainModeLbl
        MainModeLbl = Label(text="Разделение по коробам", background="white", font=("Arial", 12))
        MainModeLbl.place(x=16, y=6)







                ######## Divide widgets #######
                ################################
                
        # >>>>> Divide input file widgets block <<<<
        # ===========================================
        
        global DivideInputDirLbl
        DivideInputDirLbl = Label(text="Выберите файлы для разделения:", background="white", font=("Arial", 10))
        
        global DivideInputDirEntry
        DivideInputDirEntry = Entry(fg="black", bg="white", width=46)
        DivideInputDirEntry.configure(state = DISABLED)
        
        global DivideInputDirChooseBtn
        DivideInputDirChooseBtn = Button(text='Выбор', command=DivideInputDirChoose)
        
        global DivideInputDirPCountLbl
        DivideInputDirPCountLbl = Label(text="", background="white")
        
        
        # >>>>>>> Divide output widgets block <<<<<<
        # ===========================================
        
        global DivideOutputDirBtnLbl
        DivideOutputDirBtnLbl = Label(text="Папка для разделенных счет-фактур:", background="white", font=("Arial", 10))
        
        global DivideOutputDirBtn
        DivideOutputDirBtn = Button(text="Выбор", command=DivideOutputDirChoose)

        global DivideOutputDirEntry
        DivideOutputDirEntry = Entry(fg="black", bg="white", width=46)
        DivideOutputDirEntry.configure(state = DISABLED)
        
        
        global DivideStartDivisionBtn
        DivideStartDivisionBtn = Button(text='Выполнить разделение', command=DivideStartDivision)
        DivideStartDivisionBtn.configure(state = DISABLED)
        
        global DivideCurrentFileLbl
        DivideCurrentFileLbl = Label(text="Файл 3 из 40", background="white")

        global DivideStatusLbl
        DivideStatusLbl = Label(text="", background="white")
        
        global DivideTimeLbl
        DivideTimeLbl = Label(text="", background="white")
        
        
        
        
        
        MainModeLbl.config(text = 'Разделение по счет-фактурам:')
    
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x255+{}+{}'.format(scrnw, scrnh))
        
        
        DivideInputDirLbl.place        (x=16, y=7+40)
        DivideInputDirEntry.place      (x=20, y=30+40)
        DivideInputDirChooseBtn.place  (x=305, y=29+40, height=20)
        DivideInputDirPCountLbl.place  (x=16, y=49+40)
        DivideOutputDirBtnLbl.place    (x=17, y=87+35)
        DivideOutputDirBtn.place       (x=20, y=110+35, height=20)
        DivideOutputDirEntry.place     (x=75, y=111+35)
        DivideStartDivisionBtn.place   (x=20, y=145+37, width=150, height=25)
        DivideCurrentFileLbl.place     (x=175, y=184)
        DivideStatusLbl.place          (x=20, y=170+37)
        DivideTimeLbl.place            (x=20, y=180+44)
        
        
        
        

####### Divide functions
#########################

def DivideInputFileChoose():
    global DivideInputFile
    global DivideIsInputSel
    global DivideAllPagesCount

    DivideInputFile = filedialog.askopenfilename(title='Выберите папку на обработку', filetypes=(('PDF document', 'pdf'),))
    if DivideInputFile:
        DivideInputDirEntry.configure(state = NORMAL)
        DivideInputDirEntry.delete(0,END)
        DivideInputDirEntry.insert(0,str(DivideInputFile))
        DivideInputDirEntry.configure(state = DISABLED)
        print('Divide: IFC: DivideInputFile :', DivideInputFile)
        DivideAllPagesCount = PDFCountPages(DivideInputFile)
        DivideInputDirPCountLbl.config(text = 'Страниц в файле: ' + str(DivideAllPagesCount))
        
        DivideIsInputSel = True
        DivideOutputDirCheck()
        DivideCheckIfReady()
    else:
        print('Divide: IFC: DivideInputFile not selected')



def DivideInputDirChoose():
    global DivideInputFilesArray
    global DivideIsInputSel
    global DivideAllPagesCount

    divideinputdir = filedialog.askdirectory(title='Выберите папку с файлами на обработку')
    if divideinputdir:
        DivideInputFilesArray = []
        ScanFolder(divideinputdir, ".pdf", DivideInputFilesArray)
        
        if len(DivideInputFilesArray) > 0:
            DivideInputDirEntry.configure(state = NORMAL)
            DivideInputDirEntry.delete(0,END)
            DivideInputDirEntry.insert(0,str(Path(divideinputdir).name))
            DivideInputDirEntry.configure(state = DISABLED)
            
            print('Divide: IFC: divideinputdir :', divideinputdir)
            
            counterthread = Thread(target=DividerPageCountThread)
            counterthread.start()
            
            DivideIsInputSel = True
            DivideOutputDirCheck()
            DivideCheckIfReady()
        else:
            DivideInputDirPCountLbl.config(text = "В папке нет файлов PDF !")
            DivideIsInputSel = False
            DivideOutputDirCheck()
            DivideCheckIfReady()
    else:
        print('Divide: IFC: DivideInputFile not selected')


def DividerPageCountThread():
    global DivideInputFilesArray
    #MainModeBackBtn.configure(state = DISABLED)
    
    allpagecount = 0
    for i in range (len(DivideInputFilesArray)):
        allpagecount = allpagecount + PDFCountPages(DivideInputFilesArray[i])
    
        inputfolderinfolbl = ('Файлов в папке: {0}, считаю страницы: {1}'.format(len(DivideInputFilesArray), allpagecount))
        DivideInputDirPCountLbl.config(text = inputfolderinfolbl)
        
    inputfolderinfolbl = ('Файлов в папке: {0}, страниц: {1}'.format(len(DivideInputFilesArray), allpagecount))
    DivideInputDirPCountLbl.config(text = inputfolderinfolbl)
    #MainModeBackBtn.configure(state = NORMAL)
    
    

def DivideOutputDirChoose():
    global DivideOutputDir
    
    DivideOutputDir = filedialog.askdirectory(title='Выберите папку для счет-фактур')
    if DivideOutputDir:
        DivideOutputDirEntry.configure(state = NORMAL)
        DivideOutputDirEntry.delete(0,END)
        DivideOutputDirEntry.insert(0,str(Path(DivideOutputDir).name))
        DivideOutputDirEntry.configure(state = DISABLED)
        print('Divide: ODC: DivideOutputDir :', DivideOutputDir)
        
        DivideOutputDirCheck()
    else:
        print('Divide: ODC: DivideOutputDir not selected')


def DivideOutputDirCheck():
    global DivideInputFile
    global DivideIsInputSel
    global DivideAllPagesCount

    global DivideOutputDir
    global DivideIsOutputSel
    
    isoutdir = os.path.isdir(DivideOutputDir)
    if isoutdir:
        for k in range(5000):
            outputfile = Path (DivideOutputDir, (str(Path(DivideInputFile).name)+" - стр."+str(k)+'.pdf'))
            fileexists = os.path.isfile(outputfile)
            if fileexists:
                DivideIsOutputSel = False
                msgbxlbl = 'В этой папке уже присутствуют документы из выбранного файла! Удалите, переместите эти документы или выберите другую папку'
                messagebox.showerror("", msgbxlbl)
                break
            else:
                DivideIsOutputSel = True

    DivideCheckIfReady()
    

def DivideCheckIfReady():
    global DivideIsInputSel
    global DivideIsOutputSel
    global DivideReadyToGo
    
    print("Divide: CIR: {0}, {1}".format(DivideIsInputSel,DivideIsOutputSel))
    if DivideIsInputSel and DivideIsOutputSel:
        DivideStartDivisionBtn.configure(state = NORMAL)
        DivideReadyToGo = True
    else:
        DivideReadyToGo = False
        DivideStartDivisionBtn.configure(state = DISABLED)


def DivideBlockGUI(yes):
    if yes:
        #MainModeBackBtn.configure(state = DISABLED)
        DivideInputDirChooseBtn.configure(state = DISABLED)
        DivideOutputDirBtn.configure(state = DISABLED)
        DivideStartDivisionBtn.configure(state = DISABLED)
    else:
        #MainModeBackBtn.configure(state = NORMAL)
        DivideInputDirChooseBtn.configure(state = NORMAL)
        DivideOutputDirBtn.configure(state = NORMAL)
        DivideStartDivisionBtn.configure(state = NORMAL)




def DivideStartDivision():
    global DivideReadyToGo
    
    DivideOutputDirCheck()
    print(DivideReadyToGo)
    
    if DivideReadyToGo:
        DivideBlockGUI(True)

        minerthread = Thread(target=DivideFolderExec)
        minerthread.start()
        timethread = Thread(target=DivideTimeUpdater)
        timethread.start()
        minerthread = ""
        timethread = ""


def DivideFolderExec():
    global DivideInputFilesArray
    global DivideInputFile
    
    global DivideIsRunning
    global DivideStartedTime
    
    DivideIsRunning = True
    DivideStartedTime = time.time()
    
    for i in range (len(DivideInputFilesArray)):
        print(DivideInputFilesArray[i])
        currfilelbltxt = ("Файл {0} из {1}: {2}".format(i+1, len(DivideInputFilesArray), Path(DivideInputFilesArray[i]).name))
        DivideCurrentFileLbl.config(text = currfilelbltxt)
        DivideInputFile = DivideInputFilesArray[i]
        DivideMiner()
        
    DivideIsRunning = False
    DivideBlockGUI(False)
    
    msgbxlbl = 'Обработка файлов завершена!'
    messagebox.showinfo("", msgbxlbl)


def DivideMiner():
    global DivideInputFile
    global DivideInputPagesCount
    global DivideFirstPagesArray

    print('Divide: Miner: Started !')
    
    word = 'Счет-фактура №'
    pagecounter = 1
    finded = False
    DivideFirstPagesArray = []
    DivideFirstPagesArray.clear

    DivideInputPagesCount = PDFCountPages(DivideInputFile)
    
    pdftomine = open(DivideInputFile, 'rb')
    manager = PDFResourceManager()
    laparams = LAParams()
    dev = PDFPageAggregator(manager, laparams=laparams)
    interpreter = PDFPageInterpreter(manager, dev)
    pages = PDFPage.get_pages(pdftomine)  

    for page in pages:
        interpreter.process_page(page)
        layout = dev.get_result()
        for textbox in layout:
            if isinstance(textbox, LTText):
                for line in textbox:
                    text = line.get_text()
                    similarity = Similar(text, word)
                    if similarity > 0.9:
                        DivideFirstPagesArray.append(pagecounter)
                        print('Divide: Miner:    finded! page ' + str(pagecounter))
                        finded = True
        if finded:
            finded = False
        else:
            print('Divide: Miner:    page ' + str(pagecounter))
            
        progresslbltxt = "Поиск первых страниц... Чтение {0} из {1}, найдено: {2}".format(pagecounter, DivideInputPagesCount, len(DivideFirstPagesArray))
        DivideStatusLbl.config(text = progresslbltxt)
        pagecounter = pagecounter + 1
        
    print('Divide: Miner: Ended !')
    
    pdftomine = ''
    manager = ''
    laparams = ''
    dev = ''
    interpreter = ''
    pages = ''
    
    if len(DivideFirstPagesArray) > 0:
        DivideFileMaker()
    else:
        progresslbltxt = "Поиск завершен, счет-фактур не найдено !"
        DivideStatusLbl.config(text = progresslbltxt)
        DivideIsRunning = False
        DivideBlockGUI(False)
        msgbxlbl = 'В выбранном документе не найдено счет-фактур !'
        messagebox.showerror("", msgbxlbl)


def DivideFileMaker():
    global DivideInputFile
    global DivideInputPagesCount
    global DivideFirstPagesArray
    
    global DivideOutputDir
    
    global DivideIsRunning
    
    
    print('')
    print('Divide: FM: Started !')
    originalpdf = PyPDF2.PdfFileReader(DivideInputFile)
   
    for k in range(len(DivideFirstPagesArray)):
        if k+1 < len(DivideFirstPagesArray):
        
            print('**** Документ №: ', str(k+1))
            outputfile = Path (DivideOutputDir, (str(Path(DivideInputFile).name)+" - стр."+str(DivideFirstPagesArray[k])+'.pdf'))
            print("Номер первой стр: {0}, номер сл.первой {1}".format(DivideFirstPagesArray[k],DivideFirstPagesArray[k+1]))
            print("Итоговый файл: {0}".format(outputfile))
            print('Список страниц документа:')
            
            temparray = []
            temparray.clear()
            for x in range(DivideFirstPagesArray[k], DivideFirstPagesArray[k+1]):
                print(x)
                temparray.append(x)
            
            pdf_writer = PyPDF2.PdfFileWriter()
            for x in range(len(temparray)):
                pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))

            progresslbltxt = "Создание документов... Обработка {0} из {1}".format(k, len(DivideFirstPagesArray))
            DivideStatusLbl.config(text = progresslbltxt)

            pdf_writer.write(open(outputfile, 'wb'))
            pdf_writer = ''

            print('Обработка завершена')
            print('*******************')
            print('')


    print('**** Последний документ: ', str(len(DivideFirstPagesArray)))
    outputfile = Path (DivideOutputDir, (str(Path(DivideInputFile).name)+" - стр."+str(DivideFirstPagesArray[k])+'.pdf'))
    print("Итоговый файл: {0}".format(outputfile))
    print('Список страниц документа:')
    
    temparray.clear()
    for x in range(DivideFirstPagesArray[len(DivideFirstPagesArray)-1], DivideInputPagesCount+1):
        print(x)
        temparray.append(x)
        
    pdf_writer = PyPDF2.PdfFileWriter()
    for x in range(len(temparray)):
        pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))

    progresslbltxt = "Создание документов... Обработка {0} из {1}".format(len(DivideFirstPagesArray), len(DivideFirstPagesArray))
    DivideStatusLbl.config(text = progresslbltxt)

    pdf_writer.write(open(outputfile, 'wb'))
    pdf_writer = ''
    
    progresslbltxt = "Обработка завершена!"
    DivideStatusLbl.config(text = progresslbltxt)

    print('Обработка завершена')
    print('****************************')


    print('Documents in file: ' + str(len(DivideFirstPagesArray)))
    print('Divide: FM: Ended !')
    

    

def DivideTimeUpdater():
    global DivideIsRunning
    global DivideStartedTime

    while DivideIsRunning:
        result = time.time() - DivideStartedTime
        result = datetime2.timedelta(seconds=round(result))
        DivideTimeLbl.config(text = str(result))
        time.sleep(0.01)

#########################










def ScanFolder(folder, extension, filearray):
        filearray.clear()
        for file in os.listdir(folder):
            if file.endswith(extension):
                filearray.append(os.path.join(folder, file))

def PDFCountPages(inputfile):
    pdf = PdfFileReader(inputfile)
    pagecount = pdf.getNumPages()
    pdf = ""
    return pagecount
    
def Similar(a, b):
    return SequenceMatcher(None, a, b).ratio()
    
def ResourcePath(relative_path):    
    try:       
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

if __name__ == '__main__':
    main()
