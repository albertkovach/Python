from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
import ctypes as ct

from threading import Thread
from multiprocessing import Process, Queue, current_process, freeze_support
import psutil

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


# текстовый файл в неполном коробе




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

global DivideIsRunning
global DivideStartedTime
global DivideGUIisResized
global MaxProcessCount
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
MoveFileAmounToMove = 1000

global MoveBarcodeArray
global MoveBarcodeCount
# -------------------------



# >>>>> Sort variables <<<<
# -------------------------
global SortInputDir
global SortInputDirsArray
global SortInputDirsCount
global SortIsInputSel

global SortIsRunning
global SortStartedTime
# -------------------------



def main():
    global root
    global scrnwparam
    global scrnhparam

    root = Tk()
    root.resizable(False, False)
    
    datafile = "icon.ico"
    if not hasattr(sys, "frozen"):
        datafile = os.path.join(os.path.dirname(__file__), datafile)
    else:
        datafile = os.path.join(sys.prefix, datafile)
    root.iconbitmap(default=resource_path(datafile))
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('375x310+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()



class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")
        self.parent = parent
        self.parent.title("Docker")
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
        global DivideGUIisResized
        DivideIsInputSel = False
        DivideIsOutputSel = False
        DivideIsRunning = False
        DivideReadyToGo = False
        DivideGUIisResized = False
        
        global DivideFirstPagesArray
        DivideFirstPagesArray = []
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
        
        
        
        
        # >>> Sort var init <<<<
        # ----------------------
        global SortInputDir
        global SortInputDirsArray
        global SortInputDirsCount
        SortInputDir = ""
        SortInputDirsArray = []
        SortInputDirsCount = 0
        
        global SortIsInputSel
        global SortIsRunning
        SortIsInputSel = False
        SortIsRunning = False
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
        

        global DivideStatusLbl
        DivideStatusLbl = Label(text="", background="white")
        
        global DivideTimeLbl
        DivideTimeLbl = Label(text="", background="white")
        
        
        global DivideResizeGUIBtn
        DivideResizeGUIBtn = Button(text="➘", command=DivideResizeGUI)
        
        
        global DivideProcess1StatusLbl
        DivideProcess1StatusLbl = Label(text="Поток 1...", background="white", justify=LEFT)
        
        global DivideProcess2StatusLbl
        DivideProcess2StatusLbl = Label(text="Поток 2...", background="white", justify=LEFT)
        
        global DivideProcess3StatusLbl
        DivideProcess3StatusLbl = Label(text="Поток 3...", background="white", justify=LEFT)
        
        global DivideProcess4StatusLbl
        DivideProcess4StatusLbl = Label(text="Поток 4...", background="white", justify=LEFT)
        
        global DivideProcess5StatusLbl
        DivideProcess5StatusLbl = Label(text="Поток 5...", background="white", justify=LEFT)
        
        global DivideProcess6StatusLbl
        DivideProcess6StatusLbl = Label(text="Поток 6...", background="white", justify=LEFT)
        
        

                
                

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
        MoveBarcodeFileOpenBtn = Button(text="Создать", command=MoveBarcodeFileOpen)

        global MoveBarcodeSelEntry
        MoveBarcodeSelEntry = Entry(fg="black", bg="white", width=32)
        MoveBarcodeSelEntry.configure(state = DISABLED)
        
        global MoveBarcodeCountLbl
        MoveBarcodeCountLbl = Label(text="", background="white")
        
        global MoveRefreshBtn
        MoveRefreshBtn = Button(text='⟳', command=MoveRefresh, font=("Arial", 12))
        
        global MoveStartMovingBtn
        MoveStartMovingBtn = Button(text="Выполнить", command=MoveStartMoving)
        MoveStartMovingBtn.configure(state = DISABLED)
        
        
        
        
        
                ######### Merge widgets ########
                ################################
                
        # >>>>>>>>>>> Merge widgets block <<<<<<<<<<<
        # ===========================================
        
        global SortInputDirPathLbl
        SortInputDirPathLbl = Label(text="Выберите папку с коробами:", background="white", font=("Arial", 10))
        
        global SortInputDirEntry
        SortInputDirEntry = Entry(fg="black", bg="white", width=46)
        SortInputDirEntry.configure(state = DISABLED)
        
        global SortInputDirChooseBtn
        SortInputDirChooseBtn = Button(text='Выбор', command=SortInputDirChoose)

        global SortValidDirsCountLbl
        SortValidDirsCountLbl = Label(text="", background="white")
        
        
        global SortStartCombineBtn
        SortStartCombineBtn = Button(text='Запуск обработки', command=SortStartCombineThread)
        SortStartCombineBtn.configure(state = DISABLED)

        global SortStatus1Lbl
        SortStatus1Lbl = Label(text="", background="white")
        
        global SortStatus2Lbl
        SortStatus2Lbl = Label(text="", background="white")
        
        global SortTimeLbl
        SortTimeLbl = Label(text="", background="white")
        

        
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
        root.geometry('375x245+{}+{}'.format(scrnw, scrnh))
        
        SetModeDividePdfBtn.place(x=55, y=60+10, width = 300, height=27)
        SetModeMoveByBoxBtn.place(x=55, y=110+10, width = 300, height=27)
        SetModeSortAndMergeBtn.place(x=55, y=160+10, width = 300, height=27)

        DivideInputDirLbl.place_forget()
        DivideInputDirEntry.place_forget()
        DivideInputDirChooseBtn.place_forget()
        DivideInputDirPCountLbl.place_forget()
        DivideOutputDirBtnLbl.place_forget()
        DivideOutputDirBtn.place_forget()
        DivideOutputDirEntry.place_forget()
        DivideStartDivisionBtn.place_forget()
        DivideStatusLbl.place_forget()
        DivideTimeLbl.place_forget()
        DivideResizeGUIBtn.place_forget()
        DivideProcess1StatusLbl.place_forget()
        DivideProcess2StatusLbl.place_forget()
        DivideProcess3StatusLbl.place_forget()
        DivideProcess4StatusLbl.place_forget()
        DivideProcess5StatusLbl.place_forget()
        DivideProcess6StatusLbl.place_forget()

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
        
        SortInputDirPathLbl.place_forget()
        SortInputDirEntry.place_forget()
        SortInputDirChooseBtn.place_forget()
        SortValidDirsCountLbl.place_forget()
        SortStartCombineBtn.place_forget()
        SortStatus1Lbl.place_forget()
        SortStatus2Lbl.place_forget()
        SortTimeLbl.place_forget()

        
    if CurrentUIpage == 1:
        MainModeLbl.config(text = 'Разделение по счет-фактурам:')
        MainModeBackBtn.place(x=310, y=6, height=20)
    
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x225+{}+{}'.format(scrnw, scrnh))
        
        SetModeDividePdfBtn.place_forget()
        SetModeMoveByBoxBtn.place_forget()
        SetModeSortAndMergeBtn.place_forget()
        
        DivideInputDirLbl.place        (x=16, y=7+40)
        DivideInputDirEntry.place      (x=20, y=30+40)
        DivideInputDirChooseBtn.place  (x=305, y=29+40, height=20)
        DivideInputDirPCountLbl.place  (x=16, y=49+40)
        DivideOutputDirBtnLbl.place    (x=17, y=87+35)
        DivideOutputDirBtn.place       (x=20, y=110+35, height=20)
        DivideOutputDirEntry.place     (x=75, y=111+35)
        DivideStartDivisionBtn.place   (x=20, y=145+37, width=150, height=25)
        DivideStatusLbl.place          (x=175, y=184)
        DivideTimeLbl.place            (x=285, y=200)
        DivideResizeGUIBtn.place       (x=340, y=202, height=16)
        DivideProcess1StatusLbl.place     (x=20, y=230)
        DivideProcess2StatusLbl.place     (x=20, y=270)
        DivideProcess3StatusLbl.place     (x=20, y=310)
        DivideProcess4StatusLbl.place     (x=20, y=350)
        DivideProcess5StatusLbl.place     (x=20, y=390)
        DivideProcess6StatusLbl.place     (x=20, y=430)

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
        
        SortInputDirPathLbl.place_forget()
        SortInputDirEntry.place_forget()
        SortInputDirChooseBtn.place_forget()
        SortValidDirsCountLbl.place_forget()
        SortStartCombineBtn.place_forget()
        SortStatus1Lbl.place_forget()
        SortStatus2Lbl.place_forget()
        SortTimeLbl.place_forget()

        
    if CurrentUIpage == 2:
        MainModeLbl.config(text = 'Разделение по коробам:')
        MainModeBackBtn.place(x=310, y=6, height=20)
    
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x310+{}+{}'.format(scrnw, scrnh))
        
        SetModeDividePdfBtn.place_forget()
        SetModeMoveByBoxBtn.place_forget()
        SetModeSortAndMergeBtn.place_forget()
        
        DivideInputDirLbl.place_forget()
        DivideInputDirEntry.place_forget()
        DivideInputDirChooseBtn.place_forget()
        DivideInputDirPCountLbl.place_forget()
        DivideOutputDirBtnLbl.place_forget()
        DivideOutputDirBtn.place_forget()
        DivideOutputDirEntry.place_forget()
        DivideStartDivisionBtn.place_forget()
        DivideStatusLbl.place_forget()
        DivideTimeLbl.place_forget()
        DivideResizeGUIBtn.place_forget()
        DivideProcess1StatusLbl.place_forget()
        DivideProcess2StatusLbl.place_forget()
        DivideProcess3StatusLbl.place_forget()
        DivideProcess4StatusLbl.place_forget()
        DivideProcess5StatusLbl.place_forget()
        DivideProcess6StatusLbl.place_forget()

        MoveInputDirPathLbl.place       (x=16, y=7+40)
        MoveInputDirEntry.place         (x=20, y=30+40)
        MoveInputDirChooseBtn.place     (x=305, y=28+40, height=20)
        MoveRefreshBtn.place            (x=335, y=163+40, width = 20, height=20)
        MoveFilesCountLbl.place         (x=17, y=49+40)
        MoveOutputDirBtnLbl.place       (x=17, y=87+40)
        MoveOutputDirBtn.place          (x=20, y=110+40, height=20)
        MoveOutputDirEntry.place        (x=75, y=111+40)
        MoveBarcodeSelLbl.place         (x=17, y=141+40)
        MoveBarcodeSelBtn.place         (x=20, y=164+40, height=20)
        MoveBarcodeFileOpenBtn.place    (x=275, y=163+40, width = 55, height=20)
        MoveBarcodeSelEntry.place       (x=75, y=165+40)
        MoveBarcodeCountLbl.place       (x=17, y=185+40)
        MoveStartMovingBtn.place        (x=20, y=225+40, width=90, height=25)
        
        SortInputDirPathLbl.place_forget()
        SortInputDirEntry.place_forget()
        SortInputDirChooseBtn.place_forget()
        SortValidDirsCountLbl.place_forget()
        SortStartCombineBtn.place_forget()
        SortStatus1Lbl.place_forget()
        SortStatus2Lbl.place_forget()
        SortTimeLbl.place_forget()
        
    if CurrentUIpage == 3:
        MainModeLbl.config(text = 'Объединение счет-фактур:')
        MainModeBackBtn.place(x=310, y=6, height=20)
        
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x180+{}+{}'.format(scrnw, scrnh))
        
        SetModeDividePdfBtn.place_forget()
        SetModeMoveByBoxBtn.place_forget()
        SetModeSortAndMergeBtn.place_forget()

        DivideInputDirLbl.place_forget()
        DivideInputDirEntry.place_forget()
        DivideInputDirChooseBtn.place_forget()
        DivideInputDirPCountLbl.place_forget()
        DivideOutputDirBtnLbl.place_forget()
        DivideOutputDirBtn.place_forget()
        DivideOutputDirEntry.place_forget()
        DivideStartDivisionBtn.place_forget()
        DivideStatusLbl.place_forget()
        DivideTimeLbl.place_forget()
        DivideResizeGUIBtn.place_forget()
        DivideProcess1StatusLbl.place_forget()
        DivideProcess2StatusLbl.place_forget()
        DivideProcess3StatusLbl.place_forget()
        DivideProcess4StatusLbl.place_forget()
        DivideProcess5StatusLbl.place_forget()
        DivideProcess6StatusLbl.place_forget()

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
        
        SortInputDirPathLbl.place       (x=16, y=7+40)
        SortInputDirEntry.place         (x=20, y=30+40)
        SortInputDirChooseBtn.place     (x=305, y=29+40, height=20)
        SortValidDirsCountLbl.place     (x=17, y=49+40)
        SortStartCombineBtn.place       (x=17, y=100+40, width=120, height=25)
        SortStatus1Lbl.place            (x=17, y=115+0)
        SortStatus2Lbl.place            (x=145, y=140+10)
        SortTimeLbl.place               (x=145, y=125+10)
        

def FlushAll():
    # >>> Divide var flush <<
    # -----------------------
    global DivideInputFile
    global DivideOutputDir
    DivideInputFile = ""  
    DivideOutputDir = ""  

    global DivideIsInputSel
    global DivideIsOutputSel
    global DivideReadyToGo
    global DivideIsRunning
    global DivideGUIisResized
    DivideIsInputSel = False
    DivideIsOutputSel = False
    DivideReadyToGo = False
    DivideIsRunning = False
    DivideGUIisResized = False
    
    global DivideFirstPagesArray
    global DivideInputFilesArray
    DivideFirstPagesArray = []
    DivideInputFilesArray = []
    
    DivideInputDirPCountLbl.config(text = "")
    DivideStatusLbl.config(text = "")
    DivideTimeLbl.config(text = "")
    DivideProcess1StatusLbl.config(text = "Поток 1...")
    DivideProcess2StatusLbl.config(text = "Поток 2...")
    DivideProcess3StatusLbl.config(text = "Поток 3...")
    DivideProcess4StatusLbl.config(text = "Поток 4...")
    DivideProcess5StatusLbl.config(text = "Поток 5...")
    DivideProcess6StatusLbl.config(text = "Поток 6...")
    
    DivideInputDirEntry.configure(state = NORMAL)
    DivideInputDirEntry.delete(0,END)
    DivideInputDirEntry.configure(state = DISABLED)
    
    DivideOutputDirEntry.configure(state = NORMAL)
    DivideOutputDirEntry.delete(0,END)
    DivideOutputDirEntry.configure(state = DISABLED)
    
    DivideStartDivisionBtn.configure(state = DISABLED)
    # -----------------------
    
    
    
    # >>> Move var flush <<<<
    # -----------------------
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
    MoveBarcodeFileOpenBtn["text"] = "Создать"
    # -----------------------
    
    
    
    # >>> Sort var flush <<<<
    # -----------------------
    global SortInputDir
    global SortInputDirsArray
    global SortInputDirsCount
    global SortIsInputSel
    global SortIsRunning
    global SortStartedTime
    
    SortInputDir = ""
    SortInputDirsArray.clear()
    SortInputDirsCount = 0
    SortIsInputSel = False
    SortIsRunning = False
    SortStartedTime = ""
    
    SortValidDirsCountLbl.config(text = "")
    SortStatus1Lbl.config(text = "")
    SortStatus2Lbl.config(text = "")
    SortTimeLbl.config(text = "")
    
    SortInputDirEntry.configure(state = NORMAL)
    SortInputDirEntry.delete(0,END)
    SortInputDirEntry.configure(state = DISABLED)
    
    SortStartCombineBtn.configure(state = DISABLED)


    


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
    MainModeBackBtn.configure(state = DISABLED)
    
    allpagecount = 0
    for i in range (len(DivideInputFilesArray)):
        allpagecount = allpagecount + PDFCountPages(DivideInputFilesArray[i])
    
        inputfolderinfolbl = ('Файлов в папке: {0}, считаю страницы: {1}'.format(len(DivideInputFilesArray), allpagecount))
        DivideInputDirPCountLbl.config(text = inputfolderinfolbl)
        
    inputfolderinfolbl = ('Файлов в папке: {0}, страниц: {1}'.format(len(DivideInputFilesArray), allpagecount))
    DivideInputDirPCountLbl.config(text = inputfolderinfolbl)
    MainModeBackBtn.configure(state = NORMAL)


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




def DivideStartDivision():
    global DivideReadyToGo
    
    DivideOutputDirCheck()
    print(DivideReadyToGo)
    
    if DivideReadyToGo:
        timethread = Thread(target=DivideTimeUpdater)
        timethread.start()
        minerthread = Thread(target=DividerProcessManager)
        minerthread.start()
        minerthread = ""
        timethread = ""


def DividerProcessManager():
    global DivideInputFilesArray
    global DivideOutputDir
    global DivideIsRunning
    global DivideStartedTime
    global MaxProcessCount

    MaxProcessCount = 5 # from zero
    
    DivideBlockGUI(True)
    DivideIsRunning = True
    DivideStartedTime = time.time()
    
    # ProcessManager:
    runningprocessescount = 0
    fileresultarray = []
    dataq = Queue()

    for i in range (len(DivideInputFilesArray)):
        fileresultarray.append(2) # 2 -> awaiting, 1 -> in work, 0 -> ready
    print('ProcessManager: fileresultarray: {0}'.format(fileresultarray))

    processguinumarray = []
    processjustterminated = False
    firstloop = True
    while True:

        # Checking if data received
        if not dataq.empty():
            #### dataq: processnum, status, inputfile, message
            processresult = dataq.get()

            # Refreshing GUI with received data
            for g in range (len(processguinumarray)):
                if processresult[0] == processguinumarray[g]:
                    processguinum = g
                    break
                    
            processresultlbl = ['Документ №{0}: {1}'.format(processresult[0]+1, Path(processresult[2]).name), str(processresult[3])]
            if processguinum == 0:
                DivideProcess1StatusLbl.config(text = "\n".join(processresultlbl))
            elif processguinum == 1:
                DivideProcess2StatusLbl.config(text = "\n".join(processresultlbl))
            elif processguinum == 2:
                DivideProcess3StatusLbl.config(text = "\n".join(processresultlbl))
            elif processguinum == 3:
                DivideProcess4StatusLbl.config(text = "\n".join(processresultlbl))
            elif processguinum == 4:
                DivideProcess5StatusLbl.config(text = "\n".join(processresultlbl))
            elif processguinum == 5:
                DivideProcess6StatusLbl.config(text = "\n".join(processresultlbl))
                
                
            if processresult[1] == 0: # process terminated
                processjustterminated = True
                print('ProcessManager: reciever: process {0} terminated'.format(processresult[0]))
                
                runningprocessescount = runningprocessescount - 1
                fileresultarray[processresult[0]] = 0
                
                
                
        # Checking space for new process start
        if processjustterminated or firstloop == True:
            processjustterminated = False
            
            processestostart = MaxProcessCount - runningprocessescount
            
            # Checking amount of files that remains for exec
            filesawaiting = 0
            for r in range (len(fileresultarray)): # Scanning for awaiting files
                if fileresultarray[r] == 2:
                    filesawaiting = filesawaiting + 1
                    
            print('')
            print('')
            print('processjustterminated or firstloop')
            print('ProcessManager: runningprocessescount: {0}'.format(runningprocessescount))
            print('ProcessManager: filesawaiting: {0}'.format(filesawaiting))
            print('ProcessManager: processestostart: {0} ({1})'.format(processestostart, processestostart+1))

            if firstloop:
                firstloop = False
                if filesawaiting <= MaxProcessCount:
                    processestostart = filesawaiting - 1
                else:
                    processestostart = MaxProcessCount
                    
                if filesawaiting == 1:
                    processguinumarray.clear()
                    processguinumarray.append(0)
                    
                needtostartprocess = True
                print('ProcessManager: +++ First loop !')
            elif processestostart <= filesawaiting and  filesawaiting!=0 :
                needtostartprocess = True
                print('ProcessManager: +++ Need to start new process')
            else:
                needtostartprocess = False
            
            if needtostartprocess:
                needtostartprocess = False
                # Iterations for process start
                for s in range (processestostart+1):
                    for j in range (len(fileresultarray)):
                        if fileresultarray[j] == 2: # Scanning for awaiting files
                            fileresultarray[j] = 1 # Mark this file to "1 -> in work"
                            runningprocessescount = runningprocessescount + 1
                            creatingprocessfilenum = j
                            
                            print('ProcessManager: GUI procstart: before processguinumarray {0}'.format(processguinumarray))
                            if runningprocessescount > 1:
                                processguinumarray.clear()
                                for z in range (len(fileresultarray)):
                                    if fileresultarray[z] == 1:
                                        processguinumarray.append(z)
                            print('ProcessManager: GUI procstart: after processguinumarray {0}'.format(processguinumarray))
                            
                            break
                    #### DividerEm: (processnum, dataq, inputfile, outputdir)
                    subprocess = Process(target=DivideMiner, args=(creatingprocessfilenum, dataq, DivideInputFilesArray[creatingprocessfilenum], DivideOutputDir))
                    subprocess.start()
                    print('ProcessManager: ***** Just started new process ! {0}'.format(j))
                    print('ProcessManager: **** runningprocessescount: {0}'.format(runningprocessescount))
                    print('ProcessManager: **** fileresultarray {0}'.format(fileresultarray))
            print('')
            print('')
            
            
            
        # Checking if all is done
        filesready = 0
        for e in range (len(fileresultarray)): # Scanning for ready files
            if fileresultarray[e] == 0:
                filesready = filesready + 1
                
        
        
        if filesready == len(fileresultarray):
            DivideStatusLbl.config(text = 'Завершено! {0} из {1}'.format(filesready, len(fileresultarray)))
            messagebox.showinfo("", "Обработка завершена !")
            print('ProcessManager: All done !!!')
            break
        else:
            DivideStatusLbl.config(text = 'Выполнено {0} из {1}'.format(filesready, len(fileresultarray)))

    DivideIsRunning = False
    DivideBlockGUI(False)


def DivideMiner(processnum, dataq, DivideInputFile, DivideOutputDir):

    processname = current_process().name
    print("= Divider №{0}--{1} STARTED !: {2}".format(processnum, processname, DivideInputFile))
    
    word = 'Счет-фактура №'
    pagecounter = 1
    finded = False
    firstpagesarray = []
    firstpagesarray.clear

    pdf = PdfFileReader(DivideInputFile)
    DivideInputPagesCount = pdf.getNumPages()

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
                        firstpagesarray.append(pagecounter)
                        print("= Divider №{0}--{1}, file: {2}    finded! page {3}".format(processnum, processname, DivideInputFile, pagecounter))
                        finded = True
        if finded:
            finded = False
        else:
            print("= Divider №{0}--{1}, file: {2}    page {3}".format(processnum, processname, DivideInputFile, pagecounter))
            
        progresslbltxt = "Поиск первых страниц... Чтение {0} из {1}, найдено: {2}".format(pagecounter, DivideInputPagesCount, len(firstpagesarray))
        dataq.put([processnum, 1, DivideInputFile, progresslbltxt])
        #DivideStatusLbl.config(text = progresslbltxt)
        pagecounter = pagecounter + 1
        
    print("= Divider №{0}--{1}, ENDED !: {2}".format(processnum, processname, DivideInputFile))
    
    pdftomine = ''
    manager = ''
    laparams = ''
    dev = ''
    interpreter = ''
    pages = ''
    
    
    if len(firstpagesarray) == 0:
        progresslbltxt = "Поиск завершен, счет-фактур не найдено !"
        dataq.put([processnum, 1, DivideInputFile, progresslbltxt])
        #DivideStatusLbl.config(text = progresslbltxt)
        msgbxlbl = 'В выбранном документе не найдено счет-фактур !'
        messagebox.showerror("", msgbxlbl)
    else:
        print("= Divider №{0}--{1}, FM Started !: {2}".format(processnum, processname, DivideInputFile))
        
        originalpdf = PdfFileReader(DivideInputFile)
       
        for k in range(len(firstpagesarray)):
            if k+1 < len(firstpagesarray):
            
                print('**** Документ №: ', str(k+1))
                outputfile = Path (DivideOutputDir, (str(Path(DivideInputFile).name)+" - стр."+str(firstpagesarray[k])+'.pdf'))
                print("Номер первой стр: {0}, номер сл.первой {1}".format(firstpagesarray[k],firstpagesarray[k+1]))
                print("Итоговый файл: {0}".format(outputfile))
                print('Список страниц документа:')
                
                temparray = []
                temparray.clear()
                for x in range(firstpagesarray[k], firstpagesarray[k+1]):
                    print(x)
                    temparray.append(x)
                
                pdf_writer = PdfFileWriter()
                for x in range(len(temparray)):
                    pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))

                progresslbltxt = "Создание документов... Обработка {0} из {1}".format(k, len(firstpagesarray))
                dataq.put([processnum, 1, DivideInputFile, progresslbltxt])
                #DivideStatusLbl.config(text = progresslbltxt)

                pdf_writer.write(open(outputfile, 'wb'))
                pdf_writer = ''

                print('Обработка завершена')
                print('*******************')
                print('')


        print('**** Последний документ: ', str(len(firstpagesarray)))
        outputfile = Path (DivideOutputDir, (str(Path(DivideInputFile).name)+" - стр."+str(firstpagesarray[k])+'.pdf'))
        print("Итоговый файл: {0}".format(outputfile))
        print('Список страниц документа:')
        
        temparray.clear()
        for x in range(firstpagesarray[len(firstpagesarray)-1], DivideInputPagesCount+1):
            print(x)
            temparray.append(x)
            
        pdf_writer = PdfFileWriter()
        for x in range(len(temparray)):
            pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))

        progresslbltxt = "Создание документов... Обработка {0} из {1}".format(len(firstpagesarray), len(firstpagesarray))
        dataq.put([processnum, 1, DivideInputFile, progresslbltxt])
        #DivideStatusLbl.config(text = progresslbltxt)

        pdf_writer.write(open(outputfile, 'wb'))
        pdf_writer = ''
        
        progresslbltxt = "Обработка завершена!"
        dataq.put([processnum, 0, DivideInputFile, progresslbltxt])
        #DivideStatusLbl.config(text = progresslbltxt)

        print('Обработка завершена')
        print('****************************')

        print("= Divider №{0}--{1}, FM : {2},    Documents in file: {3}".format(processnum, processname, DivideInputFile, len(firstpagesarray)))
        print("= Divider №{0}--{1}, FM Ended !: {2}".format(processnum, processname, DivideInputFile))
    



def DivideBlockGUI(yes):
    if yes:
        MainModeBackBtn.configure(state = DISABLED)
        DivideInputDirChooseBtn.configure(state = DISABLED)
        DivideOutputDirBtn.configure(state = DISABLED)
        DivideStartDivisionBtn.configure(state = DISABLED)
    else:
        MainModeBackBtn.configure(state = NORMAL)
        DivideInputDirChooseBtn.configure(state = NORMAL)
        DivideOutputDirBtn.configure(state = NORMAL)
        DivideStartDivisionBtn.configure(state = NORMAL)


def DivideResizeGUI():
    global scrnwparam
    global scrnhparam
    global DivideGUIisResized
    
    if DivideGUIisResized:
        DivideGUIisResized = not DivideGUIisResized
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x225+{}+{}'.format(scrnw, scrnh))
        DivideResizeGUIBtn.config(text = '➘')
    else:
        DivideGUIisResized = not DivideGUIisResized
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x475+{}+{}'.format(scrnw, scrnh))
        DivideResizeGUIBtn.config(text = '➚')


def DivideTimeUpdater():
    global DivideIsRunning
    global DivideStartedTime
    
    time.sleep(0.5)
    while DivideIsRunning:
        result = time.time() - DivideStartedTime
        result = datetime2.timedelta(seconds=round(result))
        DivideTimeLbl.config(text = str(result))
        time.sleep(0.01)

#########################












########## Move functions
#########################

def MoveInputDirChoose():
    global MoveInputDir

    MoveInputDir = filedialog.askdirectory(title='Выберите папку с файлами на обработку')
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
    
    MoveOutputDir = filedialog.askdirectory(title='Выберите папку с коробами')
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
    
    MoveBarcodeFile = filedialog.askopenfilename(title='Выберите файл с штрихкодами', filetypes=(('text files', 'txt'),))
    if MoveBarcodeFile:
        MoveBarcodeFileCheck()
        MoveBarcodeFileOpenBtn["text"] = "Открыть"
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
        MoveBarcodeSelEntry.insert(0,str(Path(MoveBarcodeFile).name))
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
        MoveBarcodeFileOpenBtn["text"] = "Открыть"


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
            print('MOVER: CIR: sets:', str(set(temparray) & set(directoriesarray)))
            print('MOVER: CIR: Folders intersect:', str(intersect))
        
            if intersect:
                intersectboxesarray = []
                intersectdata = list(set(temparray) & set(directoriesarray))
                for m in range (len(intersectdata)):
                    intersectboxesarray.append(Path(intersectdata[m]).name)
                    print('MOVER: CIR:  = intersectdata', str(intersectboxesarray[m]))
                msgbxlbl = ['Некоторые короба из списка штрихкодов уже созданы !', 'Список этих коробов: {0}'.format(intersectboxesarray)]
                messagebox.showerror("", "\n".join(msgbxlbl))
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
                filesremain = MoveFileAmounToMove - len(temparray)
                uncompleted = False
            else:
                filestomove = len(temparray)
                uncompleted = True
                print("MOVER: SM: Folder: {0}, will be UNCOMPLETED".format(MoveBarcodeArray[i]))
            
            for k in range(filestomove):
                filename = Path(MoveFilesArray[k])
                MoveFilesPath = Path(MoveOutputDir, MoveBarcodeArray[i], "оригиналы", filename.name)
                RevisedMoveFilesPath = MoveFilesPath.as_posix()
                print("MOVER: SM: Folder: {0}, Num {1}, file: {2}".format(MoveBarcodeArray[i],k,str(filename.name)))
                shutil.move(MoveFilesArray[k], RevisedMoveFilesPath)
                #shutil.copyfile(MoveFilesArray[k], RevisedMoveFilesPath)
            
            if uncompleted:
                filesremain = MoveFileAmounToMove - len(temparray)
                warningfilepath = Path(MoveOutputDir, MoveBarcodeArray[i], "!!! Неполный, нет {0} файлов.txt".format(filesremain))
                fileexists = os.path.isfile(warningfilepath)
                if fileexists:
                    print("MOVER: SM: Folder: {0} WARNING ALREADY EXISTS ?????".format(MoveBarcodeArray[i]))
                else:
                    open(warningfilepath, 'w').close()

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









######## Sort functions
#########################

def SortInputDirChoose():
    global SortInputDir
    global SortInputDirsArray

    SortInputDirsArray = []

    SortInputDir = filedialog.askdirectory(title='Выберите папку с коробами')
    if SortInputDir:
        SortInputDirEntry.configure(state = NORMAL)
        SortInputDirEntry.delete(0,END)
        SortInputDirEntry.insert(0,str(SortInputDir))
        SortInputDirEntry.configure(state = DISABLED)
        
        print('Sort: IDC: MoveInputDir :', SortInputDir)
        SortInputDirCheck()
    else:
        print('Sort: IDC: MoveInputDir not selected')
        
        
def SortInputDirCheck():
    global SortInputDir
    global SortInputDirsArray
    global SortInputDirsCount
    global SortIsInputSel

    tempdirsarray = []
    tempsubdirarray = []
    tempsubdirpdfarray = []
    ScanFolder(SortInputDir, "", tempdirsarray)
    SortInputDirsArray = []
    SortInputDirsArray.clear
    SortInputDirsCount = 0
    
    if not len(tempdirsarray) > 0:
        lbltext = "Папка не содержит коробов !"
        SortValidDirsCountLbl.config(text = lbltext)
        SortIsInputSel = False
    else:
        for i in range (len(tempdirsarray)):
            ScanFolder(tempdirsarray[i], ".pdf", tempsubdirarray)
            if len(tempsubdirarray) > 0:
                print('Sort: IDC: {0} folder is NOT valid: {1}'.format(i, tempdirsarray[i]))
            else:
                tempsubdirpdfpath = Path(tempdirsarray[i], 'оригиналы')
                issubdirpdfpath = os.path.isdir(tempsubdirpdfpath)
                if issubdirpdfpath:
                    ScanFolder(tempsubdirpdfpath, ".pdf", tempsubdirpdfarray)
                    SortInputDirsArray.append(tempdirsarray[i])
                    print('Sort: IDC: {0} folder is valid, PDF: {1}, path: {2}'.format(i, len(tempsubdirpdfarray), tempdirsarray[i]))
                else:
                    print('Sort: IDC: {0} folder is NOT valid. No PDF files in orig folder, path: {1}. '.format(i, tempdirsarray[i]))

        SortInputDirsCount = len(SortInputDirsArray)
        
        if SortInputDirsCount > 0:
            lbltext = "Необработанных коробов: {0}, обработано {1}".format(SortInputDirsCount, len(tempdirsarray)-SortInputDirsCount)
            SortIsInputSel = True
            SortStartCombineBtn.configure(state = NORMAL)
            
        else:
            lbltext = "Нет коробов для обработки !"
            SortIsInputSel = False
            SortStartCombineBtn.configure(state = DISABLED)
        
        SortValidDirsCountLbl.config(text = lbltext)


def SortStartCombineThread():
    DivideBlockGUI(True)

    sortthread = Thread(target=SortStartCombine)
    sortthread.start()
    timethread = Thread(target=SortTimeUpdater)
    timethread.start()
    sortthread = ""
    timethread = ""


def SortStartCombine():
    global SortInputDirsArray
    global SortInputDirsCount
    global SortIsInputSel
    
    global SortIsRunning
    global SortStartedTime
    
    SortIsRunning = True
    SortStartedTime = time.time()
    SortBlockGUI(True)

    print('Sort: SC: started !')
    for i in range (SortInputDirsCount):
        print('Sort: SC: working in: {0}'.format(SortInputDirsArray[i]))
        stts1lbl = "Обработка короба {0}: {1} из {2}".format(Path(SortInputDirsArray[i]).name, i, len(SortInputDirsArray))
        SortStatus1Lbl.config(text = stts1lbl)
        origfilespath = Path(SortInputDirsArray[i], 'оригиналы')
        
        # Create temp folder and copy PDFs
        tempfolderpath = Path(SortInputDirsArray[i], 'временные')
        istempfolderpath = os.path.isdir(tempfolderpath)
        if istempfolderpath:
            shutil.rmtree(tempfolderpath)
            os.makedirs(tempfolderpath)
        else:
            os.makedirs(tempfolderpath)

        pdfarray = []
        ScanFolder(origfilespath, ".pdf", pdfarray)
        for m in range(len(pdfarray)):
            outputfile = Path(tempfolderpath, Path(pdfarray[m]).name).as_posix()
            shutil.copyfile(pdfarray[m], outputfile)
             
             
        # Sort by folders and convert to even
        onepagedirpath = Path(SortInputDirsArray[i], 'onepage')
        isonepagedirpath = os.path.isdir(onepagedirpath)
        if isonepagedirpath:
            shutil.rmtree(onepagedirpath)
            os.makedirs(onepagedirpath)
        else:
            os.makedirs(onepagedirpath)
            
        multipagedirpath = Path(SortInputDirsArray[i], 'multipage')
        ismultipagedirpath = os.path.isdir(multipagedirpath)
        if ismultipagedirpath:
            shutil.rmtree(multipagedirpath)
            os.makedirs(multipagedirpath)
        else:
            os.makedirs(multipagedirpath)

        pdfarray = []
        pdfile = ""
        ScanFolder(tempfolderpath, ".pdf", pdfarray)
        for k in range (len(pdfarray)):
            stts2lbl = "Сортировка файлов: {0} из {1}".format(k, len(pdfarray))
            SortStatus2Lbl.config(text = stts2lbl)
        
            pdfile = PdfFileReader(pdfarray[k])
            filepagecount = pdfile.getNumPages()
            
            if filepagecount == 1:
                outputfile = Path(onepagedirpath, Path(pdfarray[k]).name).as_posix()
                print('Sort: SC: onepage - '+str(pdfarray[k]))
                shutil.move(pdfarray[k], outputfile)
                pdfile = ""
            elif filepagecount % 2 == 1:
                _, _, w, h = pdfile.getPage(0)['/MediaBox']
                pdfile = ""
                print('Sort: SC: odd - ' + str(pdfarray[k]))
                outputfile = Path(multipagedirpath, Path(pdfarray[k]).name).as_posix()
                PDFmakeEvenPage(pdfarray[k], outputfile)
                os.remove(pdfarray[k])
            else:
                outputfile = Path(multipagedirpath, Path(pdfarray[k]).name).as_posix()
                print('Sort: SC: even - '+str(pdfarray[k]))
                shutil.move(pdfarray[k], outputfile)
                pdfile = ""
        shutil.rmtree(tempfolderpath)
        
        # Merging sorted files
        stts2lbl = "Сборка одностраничных документов"
        SortStatus2Lbl.config(text = stts2lbl)
        pdfarray = []
        pdfile = ""
        ScanFolder(onepagedirpath, ".pdf", pdfarray)
        if len(pdfarray) > 0:
            pdfmerger = PdfFileMerger()
            for h in range (len(pdfarray)):
                pdfmerger.append(pdfarray[h])
            onepagepdfpath = Path(SortInputDirsArray[i], (str(Path(SortInputDirsArray[i]).name) + " - Односторонняя печать.pdf")).as_posix()
            pdfmerger.write(onepagepdfpath)
            pdfmerger.close()
            pdfmerger = ""
        shutil.rmtree(onepagedirpath)
            
        stts2lbl = "Сборка многостраничных документов"
        SortStatus2Lbl.config(text = stts2lbl)
        pdfarray = []
        pdfile = ""
        ScanFolder(multipagedirpath, ".pdf", pdfarray)
        if len(pdfarray) > 0:
            pdfmerger = PdfFileMerger()
            for h in range (len(pdfarray)):
                pdfmerger.append(pdfarray[h])
            multipagepdfpath = Path(SortInputDirsArray[i], (str(Path(SortInputDirsArray[i]).name) + " - Двухсторонняя печать.pdf")).as_posix()
            pdfmerger.write(multipagepdfpath)
            pdfmerger.close()
            pdfmerger = ""
        shutil.rmtree(multipagedirpath)
        
        stts2lbl = "Короб завершен !"
        SortStatus2Lbl.config(text = stts2lbl)
        

    print('Sort: SC: ended !')
    SortInputDirCheck()
    
    SortIsRunning = False
    SortBlockGUI(False)
    
    stts1lbl = "Обработка коробов завершена !"
    SortStatus1Lbl.config(text = stts1lbl)
    stts2lbl = ""
    SortStatus2Lbl.config(text = stts2lbl)
    
    msgbxlbl = 'Обработка коробов завершена !'
    messagebox.showinfo("", msgbxlbl)



def SortTimeUpdater():
    global SortIsRunning
    global SortStartedTime

    while SortIsRunning:
        result = time.time() - SortStartedTime
        result = datetime2.timedelta(seconds=round(result))
        SortTimeLbl.config(text = str(result))
        time.sleep(0.01)


def SortBlockGUI(yes):
    if yes:
        MainModeBackBtn.configure(state = DISABLED)
        SortInputDirChooseBtn.configure(state = DISABLED)
        SortStartCombineBtn.configure(state = DISABLED)
    else:
        MainModeBackBtn.configure(state = NORMAL)
        SortInputDirChooseBtn.configure(state = NORMAL)
        SortStartCombineBtn.configure(state = NORMAL)



#########################














def CheckListDuplicates(listOfElems):
    if len(listOfElems) == len(set(listOfElems)):
        return False
    else:
        return True

def PDFmakeEvenPage(in_fpath, out_fpath):
    reader = PyPDF2.PdfFileReader(in_fpath)
    writer = PyPDF2.PdfFileWriter()
    for i in range(reader.getNumPages()):
        writer.addPage(reader.getPage(i))
    if reader.getNumPages() % 2 == 1:
        _, _, w, h = reader.getPage(0)['/MediaBox']
        writer.addBlankPage(w, h)
    with open(out_fpath, 'wb') as fd:
        writer.write(fd)
    fd.close()


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
    
    
def resource_path(relative_path):    
    try:       
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

if __name__ == '__main__':
    multiprocessing.freeze_support()
    main()
