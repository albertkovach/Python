from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk

from threading import Thread

import time
import datetime as datetime2
from datetime import datetime

from pathlib import Path
import shutil
import os

import PyPDF2
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter
from PyPDF2 import PdfFileMerger

from pdfminer.layout import LAParams, LTTextBox, LTText, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator

from difflib import SequenceMatcher


# печать уже существующих коробов




# >>>>> GUI variables <<<<<
# -------------------------
global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

global CurrentUIpage
# -------------------------



# >>> Divider variables <<<
# -------------------------
global DividerInputFile
global DividerIsInputSel

global DividerOutputDir
global DividerIsOutputSel

global DividerAllPagesCount
global DividerFirstPagesArray

global DividerIsRunning
global DividerStartedTime
# -------------------------



# >>>>> Move variables <<<<
# -------------------------
global MoveInputDir
global MoveAvlFolders
global MoveAvlFullFolders
global MoveAvlFolderRestFiles

global MoveOutputDir
global MoveBarcodeFile

global MoveIsInputSel
global MoveIsOutputSel
global MoveIsBarcodeSel
global MoveReadyToGo

global MoveFilesArray
global MoveValidFilesCount
global MoveFileAmounToMove
MoveFileAmounToMove = 10

global MoveBarcodeArray
global MoveBarcodeCount
# -------------------------







def main():
    global root
    global scrnwparam
    global scrnhparam

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('375x310+{}+{}'.format(scrnw, scrnh))
        
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
        # >> Divider var init <<
        # ----------------------
        global DividerInputFile
        global DividerOutputDir
        DividerInputFile = ""  
        DividerOutputDir = ""  
                
        global DividerIsInputSel
        global DividerIsOutputSel
        global DividerIsRunning
        DividerIsInputSel = False
        DividerIsOutputSel = False
        DividerIsRunning = False
        
        global DividerFirstPagesArray
        DividerFirstPagesArray = []
        # ----------------------
    
    
    
        # >>> Move var init <<<<
        # ----------------------
        global MoveIsInputSel
        global MoveIsOutputSel
        global MoveIsBarcodeSel
        global MoveReadyToGo
        MoveIsInputSel = False
        MoveIsOutputSel = False
        MoveIsBarcodeSel = False
        MoveReadyToGo = False
        
        global MoveInputDir
        global MoveOutputDir
        global MoveBarcodeFile
        MoveInputDir = ""
        MoveOutputDir = ""
        MoveBarcodeFile = ""
        
        global MoveFilesArray
        MoveFilesArray = ""
        # ----------------------
        
        
        

        s = ttk.Style()
        s.configure('.', background="white")
        


        # >>>>>>>>>>>>>>> Mode header <<<<<<<<<<<<<<<
        # ===========================================
        global MainModeLbl
        MainModeLbl = Label(text="Разделение по коробам", background="white", font=("Arial", 12))
        MainModeLbl.place(x=16, y=2)
        
        global MainModeBackBtn
        MainModeBackBtn = Button(text='Назад', command=SetModePageBack)
        
        global MainModeSeparator
        MainModeSeparator = ttk.Separator(root, orient='horizontal')
        MainModeSeparator.place(relx=0, y=32, relwidth=1, relheight=1)


        # >>>>>>>>>>>>>> Mode selector <<<<<<<<<<<<<<
        # ===========================================
        global SetModeDividePdfBtn
        SetModeDividePdfBtn = Button(text='Разделить на счет-фактуры', command=SetModeDividePdf, font=("Arial", 11))
        
        global SetModeMoveByBoxBtn
        SetModeMoveByBoxBtn = Button(text='Раскидать по коробам', command=SetModeMoveByBox, font=("Arial", 11))
        
        global SetModeSortAndMergeBtn
        SetModeSortAndMergeBtn = Button(text='Объединение в коробах', command=SetModeSortAndMerge, font=("Arial", 11))




                ######## Divider widgets #######
                ################################
                
        # >>>>> Divider input file widgets block <<<<
        # ===========================================
        
        global DividerInputFileLbl
        DividerInputFileLbl = Label(text="Выберите файл для разделения:", background="white", font=("Arial", 10))
        
        global DividerInputFileEntry
        DividerInputFileEntry = Entry(fg="black", bg="white", width=46)
        DividerInputFileEntry.configure(state = DISABLED)
        
        global DividerInputFileChooseBtn
        DividerInputFileChooseBtn = Button(text='Выбор', command=DividerInputFileChoose)
        
        global DividerInputFilePCountLbl
        DividerInputFilePCountLbl = Label(text="", background="white")
        
        
        # >>>>>>> Divider output widgets block <<<<<<
        # ===========================================
        
        global DividerOutputDirBtnLbl
        DividerOutputDirBtnLbl = Label(text="Папка для разделенных счет-фактур:", background="white", font=("Arial", 10))
        
        global DividerOutputDirBtn
        DividerOutputDirBtn = Button(text="Выбор", command=DividerOutputDirChoose)

        global DividerOutputDirEntry
        DividerOutputDirEntry = Entry(fg="black", bg="white", width=46)
        DividerOutputDirEntry.configure(state = DISABLED)
        
        
        global DividerStartDivisionBtn
        DividerStartDivisionBtn = Button(text='Выполнить разделение', command=DividerStartDivision)
        DividerStartDivisionBtn.configure(state = DISABLED)

        global DividerStatusLbl
        DividerStatusLbl = Label(text="", background="white")
        
        global DividerTimeLbl
        DividerTimeLbl = Label(text="", background="white")
        
        global DividerProgressLbl # Reserved for multiple files select
        DividerProgressLbl = Label(text="", background="white")
        

                
                

                ######### Move widgets #########
                ################################
                
        # >>>>> Move input folder widgets block <<<<<
        # ===========================================
        global MoveInputDirPathLbl
        MoveInputDirPathLbl = Label(text="Выберите папку с файлами:", background="white", font=("Arial", 10))
        
        global MoveInputDirEntry
        MoveInputDirEntry = Entry(fg="black", bg="white", width=46)
        MoveInputDirEntry.configure(state = DISABLED)
        
        global MoveInputDirChooseBtn
        MoveInputDirChooseBtn = Button(text='Выбор', command=MoveInputDirChoose)

        global MoveFilesCountLbl
        MoveFilesCountLbl = Label(text="", background="white")
        
        global MoveRefreshBtn
        MoveRefreshBtn = Button(text='Обновить', command=MoveRefresh)
        
        
        # >>>>> Move output folder widgets block <<<<
        # ===========================================
        global MoveOutputDirBtnLbl
        MoveOutputDirBtnLbl = Label(text="Папка для коробов:", background="white", font=("Arial", 10))
        
        global MoveOutputDirBtn
        MoveOutputDirBtn = Button(text="Выбор", command=MoveOutputDirChoose)

        global MoveOutputDirEntry
        MoveOutputDirEntry = Entry(fg="black", bg="white", width=46)
        MoveOutputDirEntry.configure(state = DISABLED)
        
        
        # >>>>>>>> Move barcode widgets block <<<<<<<
        # ===========================================
        global MoveBarcodeSelLbl
        MoveBarcodeSelLbl = Label(text="Файл с штрихкодами:", background="white", font=("Arial", 10))
        
        global MoveBarcodeSelBtn
        MoveBarcodeSelBtn = Button(text="Выбор", command=MoveBarcodeFileChoose)
        
        global MoveBarcodeFileOpenBtn
        MoveBarcodeFileOpenBtn = Button(text="Создать файл", command=MoveBarcodeFileOpen)

        global MoveBarcodeSelEntry
        MoveBarcodeSelEntry = Entry(fg="black", bg="white", width=46)
        MoveBarcodeSelEntry.configure(state = DISABLED)
        
        global MoveBarcodeCountLbl
        MoveBarcodeCountLbl = Label(text="", background="white")
        
        global MoveStartMovingBtn
        MoveStartMovingBtn = Button(text="Выполнить", command=MoveStartMoving)
        MoveStartMovingBtn.configure(state = DISABLED)
        
        
        
        
        
                ######### Merge widgets ########
                ################################
                
        # >>>>> Merge input folder widgets block <<<<
        # ===========================================
        
        global MergeInputDirPathLbl
        MergeInputDirPathLbl = Label(text="Выберите папку с файлами:", background="white", font=("Arial", 10))
        
        global MergeInputDirEntry
        MergeInputDirEntry = Entry(fg="black", bg="white", width=46)
        MergeInputDirEntry.configure(state = DISABLED)
        
        global MergeInputDirChooseBtn
        MergeInputDirChooseBtn = Button(text='Выбор', command=MoveInputDirChoose)

        global MergeValidDirsCountLbl
        MergeValidDirsCountLbl = Label(text="Необработанных папок: 23", background="white")
        
        
MergeInputDirPathLbl
MergeInputDirEntry
MergeInputDirChooseBtn
MergeValidDirsCountLbl
        
        
        
        global CurrentUIpage
        CurrentUIpage = 0
        UIswitcher()





def SetModeDividePdf():
    global CurrentUIpage
    CurrentUIpage = 1
    UIswitcher()

def SetModeMoveByBox():
    global CurrentUIpage
    CurrentUIpage = 2
    UIswitcher()
    
def SetModeSortAndMerge():
    global CurrentUIpage
    CurrentUIpage = 3
    UIswitcher()
    
def SetModePageBack():
    FlushAll()
    global CurrentUIpage
    CurrentUIpage = 0
    UIswitcher()

def UIswitcher():
    global scrnwparam
    global scrnhparam
    global CurrentUIpage
    global MoveRefreshBtn
    
    if CurrentUIpage == 0:
        MainModeLbl.config(text = 'Выберите режим:')
        MainModeBackBtn.place_forget()
        
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x310+{}+{}'.format(scrnw, scrnh))
        
        SetModeDividePdfBtn.place(x=55, y=60, width = 300, height=27)
        SetModeMoveByBoxBtn.place(x=55, y=110, width = 300, height=27)
        SetModeSortAndMergeBtn.place(x=55, y=160, width = 300, height=27)

        DividerInputFileLbl.place_forget()
        DividerInputFileEntry.place_forget()
        DividerInputFileChooseBtn.place_forget()
        DividerInputFilePCountLbl.place_forget()
        DividerOutputDirBtnLbl.place_forget()
        DividerOutputDirBtn.place_forget()
        DividerOutputDirEntry.place_forget()
        DividerStartDivisionBtn.place_forget()
        DividerStatusLbl.place_forget()
        DividerTimeLbl.place_forget()

        MoveInputDirPathLbl.place_forget()
        MoveInputDirEntry.place_forget()
        MoveInputDirChooseBtn.place_forget()
        MoveRefreshBtn.place_forget()
        MoveFilesCountLbl.place_forget()
        MoveOutputDirBtnLbl.place_forget()
        MoveOutputDirBtn.place_forget()
        MoveOutputDirEntry.place_forget()
        MoveBarcodeSelLbl.place_forget()
        MoveBarcodeSelBtn.place_forget()
        MoveBarcodeFileOpenBtn.place_forget()
        MoveBarcodeSelEntry.place_forget()
        MoveBarcodeCountLbl.place_forget()
        MoveStartMovingBtn.place_forget()
        
    if CurrentUIpage == 1:
        MainModeLbl.config(text = 'Разделение по счет-фактурам:')
        MainModeBackBtn.place(x=310, y=6, height=20)
    
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x255+{}+{}'.format(scrnw, scrnh))
        
        SetModeDividePdfBtn.place_forget()
        SetModeMoveByBoxBtn.place_forget()
        SetModeSortAndMergeBtn.place_forget()
        
        DividerInputFileLbl.place       (x=16, y=7+40)
        DividerInputFileEntry.place     (x=20, y=30+40)
        DividerInputFileChooseBtn.place (x=305, y=29+40, height=20)
        DividerInputFilePCountLbl.place (x=16, y=49+40)
        DividerOutputDirBtnLbl.place    (x=17, y=87+35)
        DividerOutputDirBtn.place       (x=20, y=110+35, height=20)
        DividerOutputDirEntry.place     (x=75, y=111+35)
        DividerStartDivisionBtn.place   (x=20, y=145+37, width=150, height=25)
        DividerStatusLbl.place          (x=20, y=170+37)
        DividerTimeLbl.place            (x=20, y=180+44)

        MoveInputDirPathLbl.place_forget()
        MoveInputDirEntry.place_forget()
        MoveInputDirChooseBtn.place_forget()
        MoveRefreshBtn.place_forget()
        MoveFilesCountLbl.place_forget()
        MoveOutputDirBtnLbl.place_forget()
        MoveOutputDirBtn.place_forget()
        MoveOutputDirEntry.place_forget()
        MoveBarcodeSelLbl.place_forget()
        MoveBarcodeSelBtn.place_forget()
        MoveBarcodeFileOpenBtn.place_forget()
        MoveBarcodeSelEntry.place_forget()
        MoveBarcodeCountLbl.place_forget()
        MoveStartMovingBtn.place_forget()
        
    if CurrentUIpage == 2:
        MainModeLbl.config(text = 'Разделение по коробам:')
        MainModeBackBtn.place(x=310, y=6, height=20)
    
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x310+{}+{}'.format(scrnw, scrnh))
        
        SetModeDividePdfBtn.place_forget()
        SetModeMoveByBoxBtn.place_forget()
        SetModeSortAndMergeBtn.place_forget()
        
        DividerInputFileLbl.place_forget()
        DividerInputFileEntry.place_forget()
        DividerInputFileChooseBtn.place_forget()
        DividerInputFilePCountLbl.place_forget()
        DividerOutputDirBtnLbl.place_forget()
        DividerOutputDirBtn.place_forget()
        DividerOutputDirEntry.place_forget()
        DividerStartDivisionBtn.place_forget()
        DividerStatusLbl.place_forget()
        DividerTimeLbl.place_forget()

        MoveInputDirPathLbl.place       (x=16, y=7+40)
        MoveInputDirEntry.place         (x=20, y=30+40)
        MoveInputDirChooseBtn.place     (x=305, y=29+40, height=20)
        MoveRefreshBtn.place            (x=289, y=85+40, height=20)
        MoveFilesCountLbl.place         (x=17, y=49+40)
        MoveOutputDirBtnLbl.place       (x=17, y=87+40)
        MoveOutputDirBtn.place          (x=20, y=110+40, height=20)
        MoveOutputDirEntry.place        (x=75, y=111+40)
        MoveBarcodeSelLbl.place         (x=17, y=141+40)
        MoveBarcodeSelBtn.place         (x=20, y=164+40, height=20)
        MoveBarcodeFileOpenBtn.place    (x=263, y=139+40, height=20)
        MoveBarcodeSelEntry.place       (x=75, y=165+40)
        MoveBarcodeCountLbl.place       (x=17, y=185+40)
        MoveStartMovingBtn.place        (x=20, y=225+40, width=90, height=25)
        
    if CurrentUIpage == 3:
        MainModeLbl.config(text = 'Объединение счет-фактур:')
        MainModeBackBtn.place(x=310, y=6, height=20)
        
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x310+{}+{}'.format(scrnw, scrnh))
        
        SetModeDividePdfBtn.place_forget()
        SetModeMoveByBoxBtn.place_forget()
        SetModeSortAndMergeBtn.place_forget()

        DividerInputFileLbl.place_forget()
        DividerInputFileEntry.place_forget()
        DividerInputFileChooseBtn.place_forget()
        DividerInputFilePCountLbl.place_forget()
        DividerOutputDirBtnLbl.place_forget()
        DividerOutputDirBtn.place_forget()
        DividerOutputDirEntry.place_forget()
        DividerStartDivisionBtn.place_forget()
        DividerStatusLbl.place_forget()
        DividerTimeLbl.place_forget()

        MoveInputDirPathLbl.place_forget()
        MoveInputDirEntry.place_forget()
        MoveInputDirChooseBtn.place_forget()
        MoveRefreshBtn.place_forget()
        MoveFilesCountLbl.place_forget()
        MoveOutputDirBtnLbl.place_forget()
        MoveOutputDirBtn.place_forget()
        MoveOutputDirEntry.place_forget()
        MoveBarcodeSelLbl.place_forget()
        MoveBarcodeSelBtn.place_forget()
        MoveBarcodeFileOpenBtn.place_forget()
        MoveBarcodeSelEntry.place_forget()
        MoveBarcodeCountLbl.place_forget()
        MoveStartMovingBtn.place_forget()

def FlushAll():
    global DividerInputFile
    global DividerOutputDir
    DividerInputFile = ""  
    DividerOutputDir = ""  

    global DividerIsInputSel
    global DividerIsOutputSel
    global DividerIsRunning
    DividerIsInputSel = False
    DividerIsOutputSel = False
    DividerIsRunning = False
    
    global DividerFirstPagesArray
    DividerFirstPagesArray = []
    
    DividerInputFilePCountLbl.config(text = "")
    DividerStatusLbl.config(text = "")
    DividerTimeLbl.config(text = "")
    
    DividerInputFileEntry.configure(state = NORMAL)
    DividerInputFileEntry.delete(0,END)
    DividerInputFileEntry.configure(state = DISABLED)
    
    DividerOutputDirEntry.configure(state = NORMAL)
    DividerOutputDirEntry.delete(0,END)
    DividerOutputDirEntry.configure(state = DISABLED)
    
    DividerStartDivisionBtn.configure(state = DISABLED)




    global MoveInputDir
    global MoveAvlFolders
    global MoveAvlFullFolders
    global MoveAvlFolderRestFiles
    MoveInputDir = ""
    MoveAvlFolders = 0
    MoveAvlFullFolders = 0
    MoveAvlFolderRestFiles = 0.0

    global MoveOutputDir
    global MoveBarcodeFile
    MoveOutputDir = ""
    MoveBarcodeFile = ""

    global MoveIsInputSel
    global MoveIsOutputSel
    global MoveIsBarcodeSel
    global MoveReadyToGo
    MoveIsInputSel = False
    MoveIsOutputSel = False
    MoveIsBarcodeSel = False
    MoveReadyToGo = False

    global MoveFilesArray
    global MoveValidFilesCount
    MoveFilesArray = []
    MoveFilesArray.clear
    MoveValidFilesCount = 0

    global MoveBarcodeArray
    global MoveBarcodeCount
    MoveBarcodeArray = []
    MoveBarcodeArray.clear
    MoveBarcodeCount = 0
        
    MoveFilesCountLbl.config(text = "")
    MoveInputDirEntry.configure(state = NORMAL)
    MoveInputDirEntry.delete(0,END)
    MoveInputDirEntry.configure(state = DISABLED)
    
    MoveOutputDirEntry.configure(state = NORMAL)
    MoveOutputDirEntry.delete(0,END)
    MoveOutputDirEntry.configure(state = DISABLED)
    
    MoveBarcodeCountLbl.config(text = "")
    MoveBarcodeSelEntry.configure(state = NORMAL)
    MoveBarcodeSelEntry.delete(0,END)
    MoveBarcodeSelEntry.configure(state = DISABLED)
    MoveBarcodeFileOpenBtn["text"] = "Создать файл"
    
    



    


####### Divider functions
#########################

def DividerInputFileChoose():
    global DividerInputFile
    global DividerIsInputSel
    global DividerAllPagesCount

    DividerInputFile = filedialog.askopenfilename(filetypes=(('PDF document', 'pdf'),))
    if DividerInputFile:
        DividerInputFileEntry.configure(state = NORMAL)
        DividerInputFileEntry.delete(0,END)
        DividerInputFileEntry.insert(0,str(DividerInputFile))
        DividerInputFileEntry.configure(state = DISABLED)
        print('DIVIDER: IFC: DividerInputFile :', DividerInputFile)
        DividerAllPagesCount = PDFCountPages(DividerInputFile)
        DividerInputFilePCountLbl.config(text = 'Страниц в файле: ' + str(DividerAllPagesCount))
        
        DividerIsInputSel = True
        DividerOutputDirCheck()
        DividerCheckIfReady()
    else:
        print('DIVIDER: IFC: DividerInputFile not selected')


def DividerOutputDirChoose():
    global DividerOutputDir
    global DividerIsOutputSel
    
    DividerOutputDir = filedialog.askdirectory()
    if DividerOutputDir:
        DividerOutputDirEntry.configure(state = NORMAL)
        DividerOutputDirEntry.delete(0,END)
        DividerOutputDirEntry.insert(0,str(DividerOutputDir))
        DividerOutputDirEntry.configure(state = DISABLED)
        print('DIVIDER: ODC: DividerOutputDir :', DividerOutputDir)
        
        DividerOutputDirCheck()
    else:
        print('DIVIDER: ODC: DividerOutputDir not selected')


def DividerOutputDirCheck():
    global DividerInputFile
    global DividerIsInputSel
    global DividerAllPagesCount

    global DividerOutputDir
    global DividerIsOutputSel
    
    
    for k in range(5000):
        outputfile = Path (DividerOutputDir, (str(Path(DividerInputFile).name)+" - стр."+str(k)+'.pdf'))
        fileexists = os.path.isfile(outputfile)
        if fileexists:
            DividerIsOutputSel = False
            msgbxlbl = 'В этой папке уже присутствуют документы из выбранного файла! Удалите, переместите эти документы или выберите другую папку'
            messagebox.showerror("", msgbxlbl)
            break
        else:
            DividerIsOutputSel = True

    DividerCheckIfReady()
    

def DividerCheckIfReady():
    global DividerIsInputSel
    global DividerIsOutputSel
    
    print("DIVIDER: CIR: {0}, {1}".format(DividerIsInputSel,DividerIsOutputSel))
    if DividerIsInputSel and DividerIsOutputSel:
        DividerStartDivisionBtn.configure(state = NORMAL)
        DividerReadyToGo = True
    else:
        DividerStartDivisionBtn.configure(state = DISABLED)


def DividerBlockGUI(yes):
    if yes:
        MainModeBackBtn.configure(state = DISABLED)
        DividerInputFileChooseBtn.configure(state = DISABLED)
        DividerOutputDirBtn.configure(state = DISABLED)
        DividerStartDivisionBtn.configure(state = DISABLED)
    else:
        MainModeBackBtn.configure(state = NORMAL)
        DividerInputFileChooseBtn.configure(state = NORMAL)
        DividerOutputDirBtn.configure(state = NORMAL)
        DividerStartDivisionBtn.configure(state = NORMAL)




def DividerStartDivision():

    DividerBlockGUI(True)

    minerthread = Thread(target=DividerMiner)
    minerthread.start()
    timethread = Thread(target=DividerTimeUpdater)
    timethread.start()
    minerthread = ""
    timethread = ""


def DividerMiner():
    global DividerInputFile
    global DividerAllPagesCount
    global DividerFirstPagesArray
    
    global DividerIsRunning
    global DividerStartedTime
    print('DIVIDER: Miner: Started !')
    
    word = 'Счет-фактура №'
    pagecounter = 1
    finded = False
    DividerFirstPagesArray = []
    DividerFirstPagesArray.clear

    DividerIsRunning = True
    DividerStartedTime = time.time()
    
    pdftomine = open(DividerInputFile, 'rb')
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
                        DividerFirstPagesArray.append(pagecounter)
                        print('DIVIDER: Miner:    finded! page ' + str(pagecounter))
                        finded = True
        if finded:
            finded = False
        else:
            print('DIVIDER: Miner:    page ' + str(pagecounter))
            
        progresslbltxt = "Поиск первых страниц... Чтение {0} из {1}, найдено: {2}".format(pagecounter, DividerAllPagesCount, len(DividerFirstPagesArray))
        DividerStatusLbl.config(text = progresslbltxt)
        pagecounter = pagecounter + 1
        
    print('DIVIDER: Miner: Ended !')
    
    pdftomine = ''
    manager = ''
    laparams = ''
    dev = ''
    interpreter = ''
    pages = ''
    
    if len(DividerFirstPagesArray) > 0:
        DividerFileMaker()
    else:
        progresslbltxt = "Поиск завершен, счет-фактур не найдено !"
        DividerStatusLbl.config(text = progresslbltxt)
        DividerIsRunning = False
        DividerBlockGUI(False)
        msgbxlbl = 'В выбранном документе не найдено счет-фактур !'
        messagebox.showerror("", msgbxlbl)


def DividerFileMaker():
    global DividerInputFile
    global DividerAllPagesCount
    global DividerFirstPagesArray
    
    global DividerOutputDir
    
    global DividerIsRunning
    
    
    print('')
    print('DIVIDER: FM: Started !')
    originalpdf = PyPDF2.PdfFileReader(DividerInputFile)
   
    for k in range(len(DividerFirstPagesArray)):
        if k+1 < len(DividerFirstPagesArray):
        
            print('**** Документ №: ', str(k+1))
            outputfile = Path (DividerOutputDir, (str(Path(DividerInputFile).name)+" - стр."+str(DividerFirstPagesArray[k])+'.pdf'))
            print("Номер первой стр: {0}, номер сл.первой {1}".format(DividerFirstPagesArray[k],DividerFirstPagesArray[k+1]))
            print("Итоговый файл: {0}".format(outputfile))
            print('Список страниц документа:')
            
            temparray = []
            temparray.clear()
            for x in range(DividerFirstPagesArray[k], DividerFirstPagesArray[k+1]):
                print(x)
                temparray.append(x)
            
            pdf_writer = PyPDF2.PdfFileWriter()
            for x in range(len(temparray)):
                pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))

            progresslbltxt = "Создание документов... Обработка {0} из {1}".format(k, len(DividerFirstPagesArray))
            DividerStatusLbl.config(text = progresslbltxt)

            pdf_writer.write(open(outputfile, 'wb'))
            pdf_writer = ''

            print('Обработка завершена')
            print('*******************')
            print('')


    print('**** Последний документ: ', str(len(DividerFirstPagesArray)))
    outputfile = Path (DividerOutputDir, (str(Path(DividerInputFile).name)+" - стр."+str(DividerFirstPagesArray[k])+'.pdf'))
    print("Итоговый файл: {0}".format(outputfile))
    print('Список страниц документа:')
    
    temparray.clear()
    for x in range(DividerFirstPagesArray[len(DividerFirstPagesArray)-1], DividerAllPagesCount+1):
        print(x)
        temparray.append(x)
        
    pdf_writer = PyPDF2.PdfFileWriter()
    for x in range(len(temparray)):
        pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))

    progresslbltxt = "Создание документов... Обработка {0} из {1}".format(len(DividerFirstPagesArray), len(DividerFirstPagesArray))
    DividerStatusLbl.config(text = progresslbltxt)

    pdf_writer.write(open(outputfile, 'wb'))
    pdf_writer = ''
    
    progresslbltxt = "Обработка файла завершена! Извлечено документов: {0}".format(len(DividerFirstPagesArray))
    DividerStatusLbl.config(text = progresslbltxt)

    print('Обработка завершена')
    print('****************************')

    DividerIsRunning = False
    DividerBlockGUI(False)
    
    print('Documents in file: ' + str(len(DividerFirstPagesArray)))
    print('DIVIDER: FM: Ended !')
    
    msgbxlbl = 'Обработка файла завершена!'
    messagebox.showinfo("", msgbxlbl)
    

def DividerTimeUpdater():
    global DividerIsRunning
    global DividerStartedTime

    while DividerIsRunning:
        result = time.time() - DividerStartedTime
        result = datetime2.timedelta(seconds=round(result))
        DividerTimeLbl.config(text = str(result))
        time.sleep(0.01)

#########################





########## Move functions
#########################

def MoveInputDirChoose():
    global MoveInputDir

    MoveInputDir = filedialog.askdirectory()
    if MoveInputDir:
        MoveInputDirEntry.configure(state = NORMAL)
        MoveInputDirEntry.delete(0,END)
        MoveInputDirEntry.insert(0,str(MoveInputDir))
        MoveInputDirEntry.configure(state = DISABLED)
        print('MOVER: IDC: MoveInputDir :', MoveInputDir)
        MoveInputDirCheck()
    else:
        print('MOVER: IDC: MoveInputDir not selected')


def MoveOutputDirChoose():
    global MoveOutputDir
    global MoveIsOutputSel
    
    MoveOutputDir = filedialog.askdirectory()
    if MoveOutputDir:
        MoveOutputDirEntry.configure(state = NORMAL)
        MoveOutputDirEntry.delete(0,END)
        MoveOutputDirEntry.insert(0,str(MoveOutputDir))
        MoveOutputDirEntry.configure(state = DISABLED)
        print('MOVER: ODC: MoveOutputDir :', MoveOutputDir)
        
        MoveIsOutputSel = True
        MoveCheckIfReady()
    else:
        print('MOVER: ODC: MoveOutputDir not selected')


def MoveBarcodeFileChoose():
    global MoveBarcodeFile
    
    MoveBarcodeFile = filedialog.askopenfilename(filetypes=(('text files', 'txt'),))
    if MoveBarcodeFile:
        MoveBarcodeFileCheck()
        MoveBarcodeFileOpenBtn["text"] = "Открыть файл"
    else:
        print('MOVER: BFC: MoveBarcodeFile not selected')



def MoveInputDirCheck():
    global MoveInputDir
    global MoveFileAmounToMove
    global MoveIsInputSel
    
    global MoveFilesArray
    global MoveValidFilesCount
    
    global MoveAvlFolders
    global MoveAvlFullFolders
    global MoveAvlFolderRestFiles
    
    
    direxists = os.path.isdir(MoveInputDir)
    if direxists:
        MoveInputDirEntry.configure(state = NORMAL)
        MoveInputDirEntry.delete(0,END)
        MoveInputDirEntry.insert(0,str(MoveInputDir))
        MoveInputDirEntry.configure(state = DISABLED)
        
        MoveFilesArray = []
        ScanFolder(MoveInputDir, ".pdf", MoveFilesArray)
        MoveValidFilesCount = len(MoveFilesArray)

        if MoveValidFilesCount > 0:
            MoveAvlFolders = 0
            MoveAvlFullFolders = len(MoveFilesArray)/MoveFileAmounToMove
            print('MOVER: IDChk: MoveAvlFullFolders:', str(MoveAvlFullFolders))
            
            if MoveAvlFullFolders - int(MoveAvlFullFolders) > 0:
                print('MOVER: IDChk: MoveAvlFolderRestFiles: ', str(MoveAvlFullFolders - int(MoveAvlFullFolders)))
                MoveAvlFolderRestFiles = True
                MoveAvlFolders = int(MoveAvlFullFolders + 1)
                lbltext = ("Файлов PDF: {0}, нужно {1} кодов, последний - неполный".format(len(MoveFilesArray),MoveAvlFolders))
            else:
                MoveAvlFolders = int(MoveAvlFullFolders)
                lbltext = ("Файлов PDF: {0}, нужно {1} кодов".format(len(MoveFilesArray),MoveAvlFolders))
            MoveAvlFullFolders = int(MoveAvlFullFolders)
            print('MOVER: IDChk: MoveAvlFolders:', str(MoveAvlFolders))
            
        
            MoveFilesCountLbl.config(text = lbltext)
            print('MOVER: IDChk: Number of valid files -', str(MoveValidFilesCount))
            MoveIsInputSel = True
            MoveCheckIfReady()
        else:
            MoveFilesCountLbl.config(text = 'Нет файлов PDF !')
            MoveIsInputSel = False
            MoveCheckIfReady()
    else:
        MoveInputDir = ""
        MoveInputDirEntry.configure(state = NORMAL)
        MoveInputDirEntry.delete(0,END)
        MoveInputDirEntry.configure(state = DISABLED)
        MoveFileAmounToMove = ""
        MoveFilesArray.clear
        MoveValidFilesCount = ""
        MoveAvlFolders = 0
        MoveAvlFullFolders = 0
        MoveAvlFolderRestFiles = False
        
        MoveFilesCountLbl.config(text = 'Папка не существует !')
        MoveIsInputSel = False
        MoveCheckIfReady()


def MoveOutputDirCheck():
    global MoveOutputDir
    global MoveIsOutputSel
    
    direxists = os.path.isdir(MoveOutputDir)
    if not direxists:
        MoveOutputDir = ""
        MoveOutputDirEntry.configure(state = NORMAL)
        MoveOutputDirEntry.delete(0,END)
        MoveOutputDirEntry.configure(state = DISABLED)
        print('MOVER: ODChk: MoveOutputDir doesnt exist')
        MoveIsOutputSel = False
        MoveCheckIfReady()


def MoveBarcodeFileCheck():
    global MoveBarcodeFile
    global MoveBarcodeArray
    global MoveIsBarcodeSel
    global MoveAvlFolders
    global MoveIsInputSel
    

    fileexists = os.path.isfile(MoveBarcodeFile)
    if fileexists:
        MoveBarcodeSelEntry.configure(state = NORMAL)
        MoveBarcodeSelEntry.delete(0,END)
        MoveBarcodeSelEntry.insert(0,str(MoveBarcodeFile))
        MoveBarcodeSelEntry.configure(state = DISABLED)
        print('MOVER: BFChk: MoveBarcodeFile :', MoveBarcodeFile)
        
        file = open(MoveBarcodeFile,'r')
        MoveBarcodeArray = []
        MoveBarcodeArray.clear()
        for line in file:
            MoveBarcodeArray.append(line.strip()) # We don't want newlines in our list
        file.close()
        if len(MoveBarcodeArray) > 0:
            result = CheckListDuplicates(MoveBarcodeArray)
            
            if result:
                MoveBarcodeCountLbl.config(text = 'Есть повторяющиеся коды !')
                print('MOVER: BFChk: MoveBarcodeFile contain duplicates !')
                MoveIsBarcodeSel = False
                MoveCheckIfReady()
            else:
                print('MOVER: BFChk: MoveIsInputSel :', MoveIsInputSel)
                if MoveIsInputSel:
                    MoveBarcodeFileRecount()
                else:
                    MoveBarcodeCountLbl.config(text = 'Кодов в файле: ' + str(len(MoveBarcodeArray)))
                for i in range(len(MoveBarcodeArray)):
                    print('MOVER: BFChk:   - bar:', str(MoveBarcodeArray[i]))
                    
                MoveIsBarcodeSel = True
                MoveCheckIfReady()
        else:
            MoveBarcodeCountLbl.config(text = 'Штрихкодов не найдено!')
            print('MOVER: BFChk: MoveBarcodeFile is empty !')
            MoveIsBarcodeSel = False
            MoveCheckIfReady()
    else:
        MoveBarcodeFile = ""
        MoveBarcodeSelEntry.configure(state = NORMAL)
        MoveBarcodeSelEntry.delete(0,END)
        MoveBarcodeSelEntry.configure(state = DISABLED)
        
        MoveBarcodeArray.clear()
        print('MOVER: BFChk: MoveBarcodeFile doesnt exist !')
        MoveIsBarcodeSel = False
        MoveCheckIfReady()


def MoveBarcodeFileRecount():
    global MoveBarcodeArray
    global MoveAvlFolders
    
    restcodes = len(MoveBarcodeArray) - MoveAvlFolders
    if restcodes > 0:
        countlbl = ("Кодов в файле: {0}, лишних кодов: {1} ".format(len(MoveBarcodeArray),restcodes))
    if restcodes == 0:
        countlbl = ("Кодов в файле: {0}. Лишних нет !".format(len(MoveBarcodeArray)))
    if restcodes < 0:
        restcodes = restcodes * (-1)
        countlbl = ("Кодов в файле: {0}, не хватило еще {1} кодов".format(len(MoveBarcodeArray),restcodes))
    MoveBarcodeCountLbl.config(text = countlbl)


def MoveBarcodeFileOpen():
    global MoveBarcodeFile
    global MoveIsBarcodeSel
    
    fileexists = os.path.isfile(MoveBarcodeFile)
    if fileexists:
        os.startfile(MoveBarcodeFile)
    else:
        parent = Path(__file__).resolve().parent
        MoveBarcodeFile = Path(parent, "barcodes.txt")
        fileexists = os.path.isfile(MoveBarcodeFile)
        if not fileexists:
            open(MoveBarcodeFile, 'w').close()
        os.startfile(MoveBarcodeFile)
        
        MoveBarcodeFileCheck()
        MoveBarcodeFileOpenBtn["text"] = "Открыть файл"


def MoveRefresh():
    try:
        MoveInputDirCheck()
    except:
        print("MOVER: Refresh: Error refreshing InputDirCheck")

    try:
        MoveOutputDirCheck()
    except:
        print("MOVER: Refresh: Error refreshing OutputDirCheck") 

    try:
        MoveBarcodeFileCheck()
    except:
        print("MOVER: Refresh: Error refreshing BarcodeFileCheck")




def MoveCheckIfReady():
    global MoveOutputDir
    global MoveBarcodeArray
    global MoveAvlFolders
    
    global MoveIsInputSel
    global MoveIsOutputSel
    global MoveIsBarcodeSel
    global MoveReadyToGo
    
    
    print("MOVER: CIR: {0}, {1}, {2}".format(MoveIsInputSel,MoveIsOutputSel,MoveIsBarcodeSel))
    if MoveIsInputSel and MoveIsOutputSel and MoveIsBarcodeSel:
    
        # Checking existing folders from savedir and barcodearray
        allfilesarray = []
        directoriesarray = []
        temparray = []
        ScanFolder(MoveOutputDir, "", allfilesarray)
        
        for i in range(len(allfilesarray)):
            if os.path.isdir(allfilesarray[i]):
                directoriesarray.append(str(Path(allfilesarray[i]).as_posix()))

        if len(directoriesarray) > 0:
            print('MOVER: CIR: ==== Present folders:', str(len(directoriesarray)))
            for i in range(len(directoriesarray)):
                ScanFolder(directoriesarray[i], ".pdf", temparray)
                print("MOVER: CIR:  - PDFs: {0}, Folder {1}".format(len(temparray),directoriesarray[i]))
     
            temparray.clear
            temparray = MoveBarcodeArray
            print('MOVER: CIR: ==== Target folders:', str(len(temparray)))
            for i in range(len(temparray)):
                temparray[i] = Path(MoveOutputDir, temparray[i]).as_posix()
                print('MOVER: CIR:  = Folder', str(temparray[i]))

            intersect = bool(set(temparray) & set(directoriesarray))
            print('MOVER: CIR: Folders intersect:', str(intersect))
        
            if intersect:
                messagebox.showerror("", "Некоторые короба из списка штрихкодов уже созданы ! Переместите их или загрузите другой список")
                MoveIsBarcodeSel = False
                MoveBarcodeArray.clear
                MoveBarcodeCountLbl.config(text = 'По некоторым штрихкодам уже созданы короба !')
                MoveStartMovingBtn.configure(state = DISABLED)
                MoveReadyToGo = False
            else:
                # Folders exists in savedir, but not from barcodesarray
                MoveStartMovingBtn.configure(state = NORMAL)
                MoveReadyToGo = True
                
        # All good, checking how many barcodes remains after work
        else:
            MoveBarcodeFileRecount()
            MoveStartMovingBtn.configure(state = NORMAL)
            MoveReadyToGo = True

            
    else:
        MoveStartMovingBtn.configure(state = DISABLED)
        MoveReadyToGo = False

 
def MoveStartMoving():
    global MoveFilesArray
    global MoveValidFilesCount
    global MoveFileAmounToMove
    
    global MoveOutputDir
    
    global MoveBarcodeFile
    global MoveBarcodeArray
    newbarcodearray = []
    
    global MoveReadyToGo
    global MoveIsBarcodeSel
    
    global MoveAvlFolders
    folderstocreate = 0
    

    MoveRefresh()
    MoveCheckIfReady()
    
    if MoveReadyToGo:
    
        MainModeBackBtn.configure(state = DISABLED)
        MoveInputDirChooseBtn.configure(state = DISABLED)
        MoveRefreshBtn.configure(state = DISABLED)
        MoveOutputDirBtn.configure(state = DISABLED)
        MoveBarcodeSelBtn.configure(state = DISABLED)
        MoveBarcodeFileOpenBtn.configure(state = DISABLED)
        MoveStartMovingBtn.configure(state = DISABLED)
    
        if MoveAvlFolders < len(MoveBarcodeArray):
            # Codes will remain uncompleted
            print('MOVER: SM: New BarcodeFile content')
            folderstocreate = MoveAvlFolders
            
            restcodes = len(MoveBarcodeArray) - MoveAvlFolders
            barfile = open(MoveBarcodeFile, 'w').close()
            barfile = open(MoveBarcodeFile, 'a')
            
            for i in range (restcodes):
                newbarcodearray.append(MoveBarcodeArray[i+MoveAvlFolders])
                newbarcodeline = (str(newbarcodearray[i])+'\n')
                barfile.write(newbarcodeline)
                print('MOVER: SM:  - new bar', str(newbarcodearray[i]))
                
            barfile.close()
            newbarcodearray.clear
        else:
            folderstocreate = len(MoveBarcodeArray)
        
        
        for i in range(folderstocreate):
            MoveOutputPath = Path(MoveOutputDir, MoveBarcodeArray[i], "оригиналы")
            ismovedirexist = os.path.exists(MoveOutputPath)
            print('MOVER: SM:    ==== New folder:', str(MoveBarcodeArray[i]))
            
            if not ismovedirexist:
                os.makedirs(MoveOutputPath)

            # Check if is last uncompleted folder
            temparray = []
            ScanFolder(MoveInputDir, ".pdf", temparray)
            if MoveFileAmounToMove < len(temparray):
                filestomove = MoveFileAmounToMove
            else:
                filestomove = len(temparray)
                
            for k in range(filestomove):
                filename = Path(MoveFilesArray[k])
                MoveFilesPath = Path(MoveOutputDir, MoveBarcodeArray[i], "оригиналы", filename.name)
                RevisedMoveFilesPath = MoveFilesPath.as_posix()
                print("MOVER: SM: Folder: {0}, Num {1}, file: {2}".format(MoveBarcodeArray[i],k,str(filename.name)))
                shutil.move(MoveFilesArray[k], RevisedMoveFilesPath)
                #shutil.copyfile(MoveFilesArray[k], RevisedMoveFilesPath)
        
            MoveFilesArray.clear()
        
            ScanFolder(MoveInputDir, ".pdf", MoveFilesArray)
            MoveValidFilesCount = len(MoveFilesArray)
            MoveFilesCountLbl.config(text = 'Количество файлов PDF: ' + str(len(MoveFilesArray)))
            
        
        # Making report
        ScanFolder(MoveInputDir, ".pdf", MoveFilesArray)
        MoveValidFilesCount = len(MoveFilesArray)
        filesremain = 0
        if MoveValidFilesCount == 0:
            for i in range(folderstocreate):
                MoveOutputPath = Path(MoveOutputDir, MoveBarcodeArray[i], "оригиналы")
                ScanFolder(MoveOutputPath, ".pdf", temparray)
                if len(temparray)<MoveFileAmounToMove:
                    filesremain = MoveFileAmounToMove - len(temparray)
                    msgbxlbl = ['Перемещение выполнено !', 'В коробе {0} осталось место для {1} файлов'.format(MoveBarcodeArray[i], filesremain)]
                    break
            if filesremain == 0 :
                msgbxlbl = ['Перемещение выполнено !', 'Все файлы обработаны']
        if MoveValidFilesCount > 0:
            msgbxlbl = ['Перемещение выполнено !', 'Осталось файлов: {0}'.format(MoveValidFilesCount)]
        messagebox.showinfo("", "\n".join(msgbxlbl))
        
        # Ending work
        if not MoveAvlFolders < len(MoveBarcodeArray):
            open(MoveBarcodeFile, 'w').close()
            MoveBarcodeArray.clear()
            MoveIsBarcodeSel = False

        MoveBarcodeSelEntry.configure(state = NORMAL)
        MoveBarcodeSelEntry.delete(0,END)
        MoveBarcodeSelEntry.configure(state = DISABLED)
        
        MoveBarcodeCountLbl.config(text = '')
        
        MainModeBackBtn.configure(state = NORMAL)
        MoveInputDirChooseBtn.configure(state = NORMAL)
        MoveRefreshBtn.configure(state = NORMAL)
        MoveOutputDirBtn.configure(state = NORMAL)
        MoveBarcodeSelBtn.configure(state = NORMAL)
        MoveBarcodeFileOpenBtn.configure(state = NORMAL)
        
        MoveStartMovingBtn.configure(state = DISABLED)
        
        MoveRefresh()

#########################








def CheckListDuplicates(listOfElems):
    if len(listOfElems) == len(set(listOfElems)):
        return False
    else:
        return True

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

if __name__ == '__main__':
    main()
