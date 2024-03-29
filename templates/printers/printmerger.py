from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from PyPDF2 import PdfFileMerger
from threading import Thread
from datetime import datetime
from pathlib import Path
import datetime as datetime2
import shutil
import time
import os
import re
global root

import win32api
import win32print

# Albert Kovach
# 15/03/2022
# ORB Mara, cold spring, 
#       second week of war
#       end of COVID in Moscow
#       yet first wall is in work

# 23/03/2022
# added enum merged in txt
# dismissals awaiting

# 30/03/2022
# added file move
# friends dismissed


# Input folder data
global inputdir
global InputPath


# Output folder data
global savedir
global OutputPath
global OutputPathFilesEnum

global savename
global outputfile
global outputfilesenum

global ValidFilesCount
global FilesArray


# Threads
global mergethread
global timethread


# Merge thread data
global MergeMode
global MergeInProgress
global StartedMergeTime

global MergeDivisor
global MergeQuotient
global MergeRemainder

global CreateDocTime
global CreateFilesEnum
global MoveCompletedFiles

MergeInProgress = False


global DefDocs
global DefParts
global SettAddTime
DefDocs = 20
DefParts = 2
SettAddTime = False



def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = root.winfo_screenwidth()
    scrnh = root.winfo_screenheight()
    scrnw = scrnw//2
    scrnh = scrnh//2
    scrnw = scrnw - 185
    scrnh = scrnh - 150
    root.geometry('375x310+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()

class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.parent.title("Merge PDF")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        global mergethread
        global MergeMode
        global MergeDivisor
        global MergeQuotient
        global MergeRemainder
        global CreateFilesEnum
        global MoveCompletedFiles
        
        # Variables initialization
        MergeMode = IntVar()
        MergeMode.set(0)

        MergeDivisor = IntVar()
        MergeQuotient = 0
        MergeRemainder = 0
        
        CreateFilesEnum = IntVar()
        CreateFilesEnum.set(1)        
        MoveCompletedFiles = IntVar()
        MoveCompletedFiles.set(1)

        
        # >>>>>>>> Input folder widgets block  <<<<<<<<
        # =============================================
        global InputDirPathLbl
        InputDirPathLbl = Label(text="Выберите папку с файлами:", background="white")
        InputDirPathLbl.place(x=16, y=10)
        
        global InputDirEntry
        InputDirEntry = Entry(fg="black", bg="white", width=46)
        InputDirEntry.place(x=20, y=30)
        
        global InputDirChooseBtn
        InputDirChooseBtn = Button(text='Выбор', command=InputDirChoose)
        InputDirChooseBtn.place(x=305, y=29, height=20)

        global DirCalcBtn
        DirCalcBtn = Button(text='Посчитать', command=CountFiles)
        DirCalcBtn.place(x=20, y=52, height=20)

        global FilesCountLbl
        FilesCountLbl = Label(text="", background="white")
        FilesCountLbl.place(x=95, y=52)
        # =============================================



        # >>>>>>>> Output folder widgets block <<<<<<<<
        # =============================================
        global SaveNameLbl
        SaveNameLbl = Label(text="Введите имя итогового файла:", background="white")
        SaveNameLbl.place(x=16, y=87)

        global SaveNameEntry
        SaveNameEntry = Entry(fg="black", bg="white", width=26)
        SaveNameEntry.place(x=196, y=88)

        global SaveDirBtn
        SaveDirBtn = Button(text="Выбрать путь", command=SaveDirChoose)
        SaveDirBtn.place(x=20, y=110, height=20)

        global SaveDirEntry
        SaveDirEntry = Entry(fg="black", bg="white", width=40)
        SaveDirEntry.place(x=113, y=111)
        # ==============================================



        # >>>>>>> Merging settings widgets block <<<<<<<
        # ==============================================
        global ModeSaveEnum
        ModeSaveEnum = Checkbutton(text="Записать имена в .txt", background="white", variable=CreateFilesEnum, onvalue=1, offvalue=0)
        
        global ModeMoveCompleted
        ModeMoveCompleted = Checkbutton(text="Переместить выполненные", background="white", variable=MoveCompletedFiles, onvalue=1, offvalue=0)
        
        global ModeMergeAll
        ModeMergeAll = Radiobutton(text="В один файл", background="white", variable=MergeMode, value=0, command=RadioBtnHandler)
        
        global ModeMergeByC
        ModeMergeByC = Radiobutton(text="По заданным частям", background="white", variable=MergeMode, value=1, command=RadioBtnHandler)
        
        global ModeMergeByE
        ModeMergeByE = Radiobutton(text="По равным частям", background="white", variable=MergeMode, value=2, command=RadioBtnHandler)

        global MergeDivisorLbl
        MergeDivisorLbl = Label(text="Количество частей:", background="white")
        
        global MergeDivisorSpin
        MergeDivisorSpin = Spinbox(from_=2, to=5000, textvariable=MergeDivisor, justify=CENTER, command=CountFiles)

        global MergeQuotientLbl
        MergeQuotientLbl = Label(text="Документов в одной: 0000", background="white")
        
        global MergeRemainderLbl
        MergeRemainderLbl = Label(text="Документов в последней: 0000", background="white")
        # ==============================================



        ay = 85
        global MergeBtn
        MergeBtn = Button(text="Выполнить сборку", command=StartMergingThread)
        MergeBtn.place(x=20, y=ay+175, height=20)

        global MergeStatusLbl
        MergeStatusLbl = Label(text="", background="white")
        MergeStatusLbl.place(x=145, y=ay+175)
        
        global MergeTimeLbl
        MergeTimeLbl = Label(text="", background="white")
        MergeTimeLbl.place(x=145, y=ay+192)
        
        global MergeProgressLbl
        MergeProgressLbl = Label(text="", background="white")
        MergeProgressLbl.place(x=190, y=ay+192)



def InputDirChoose(): 
    global inputdir
    
    inputdir = filedialog.askdirectory(title="Выбрать папку")
    print('======')
    if inputdir:
        InputDirEntry.delete(0,END)
        InputDirEntry.insert(0,inputdir)
        print('DChs: inputdir -', inputdir)
        CountFiles()
        RadioBtnStates()
    else:
        print('IDChs: inputdir is NOT defined')

def SaveDirChoose():
    global savedir
    print('======')
    savedir = filedialog.askdirectory(title="Выбрать папку")
    if savedir:
        SaveDirEntry.delete(0,END)
        SaveDirEntry.insert(0,savedir)
        print('SDChs: savedir -', savedir)
    else:
        print('SDChs: savedir is NOT defined')



def RadioBtnHandler():
    RadioBtnStates()
    CountFiles()

def RadioBtnStates():
    global MergeMode
    global DefParts
    global DefDocs
    
    print('======')
    print('RBH: MergeMode -', str(MergeMode.get())) 
    
    if CheckInputDir():
        ModeSaveEnum.place(x=177, y=138)
        ModeMoveCompleted.place(x=177, y=158)
        ModeMergeAll.place(x=25, y=138)
        ModeMergeByC.place(x=25, y=158)
        ModeMergeByE.place(x=25, y=178)
    
        if MergeMode.get() == 0:
            MergeDivisorLbl.place_forget()
            MergeDivisor = 1
            MergeDivisorSpin.place_forget()
            MergeQuotientLbl.place_forget()
            MergeRemainderLbl.place_forget()
        if MergeMode.get() == 1:
            MergeDivisorLbl.place(x=180, y=182)
            MergeDivisorLbl.config(text = 'Документов в части:')
            MergeDivisor = DefDocs
            MergeDivisorSpin.delete(0, END)
            MergeDivisorSpin.insert(0, MergeDivisor)
            MergeDivisorSpin.place(x=300, y=184, width=40)
            MergeQuotientLbl.place(x=180, y=203)
            MergeRemainderLbl.place(x=180, y=223)
        if MergeMode.get() == 2:
            MergeDivisorLbl.place(x=180, y=182)
            MergeDivisorLbl.config(text = 'Количество частей:')
            MergeDivisor = DefParts
            MergeDivisorSpin.delete(0, END)
            MergeDivisorSpin.insert(0, MergeDivisor)
            MergeDivisorSpin.place(x=300, y=184, width=40)
            MergeQuotientLbl.place(x=180, y=203)
            MergeRemainderLbl.place(x=180, y=223)
    else:
        ModeSaveEnum.place_forget()
        ModeMergeAll.place_forget()
        ModeMergeByC.place_forget()
        ModeMergeByE.place_forget()



def CheckInputDir():
    inputdir = InputDirEntry.get()
    if inputdir:
        print('inputdir exists')
        isinputdir = os.path.isdir(inputdir)
        print('inputdir is valid -', str(isinputdir)) 
        return isinputdir
    else:
        return False
        print('inputdir is empty')

def DateTimeUpdate(forced):
    global CreateDocTime
    
    now = datetime.now()

    if now.day < 10:
        tday = ('0' + str(now.day))
    else:
        tday = str(now.day)

    if now.month < 10:
        tmonth = ('0' + str(now.month))
    else:
        tmonth = str(now.month)

    if now.hour < 10:
        thour = ('0' + str(now.hour))
    else:
        thour = str(now.hour)

    if now.minute < 10:
        tminute = ('0' + str(now.minute))
    else:
        tminute = str(now.minute)
        
    datetimestr = (tday + "." + tmonth + "." + str(now.year) + " " + thour + "." + tminute)
    
    if forced:
        CreateDocTime = datetimestr

def SavenameGenerate(Nameless, AddDate, Name, Part):
    global CreateDocTime
    
    if Nameless:
        if AddDate:
            if Part < 1:
                savename = ("Merged "+ CreateDocTime)
            else:
                savename = ("Merged, part " + str(Part) + " - " + CreateDocTime)
        else:
            if Part < 1:
                savename = ("Merged")
            else:
                savename = ("Merged, part " + str(Part))
    else:
        if AddDate:
            if Part < 1:
                savename = (Name + " " + CreateDocTime)
            else:
                savename = (Name + " " + str(Part) + " - " + CreateDocTime)
        else:
            if Part < 1:
                savename = (Name)
            else:
                savename = (Name + ", part " + str(Part))

    return savename

def BlockGUI(block):
    if block:
        InputDirPathLbl.configure(state = DISABLED)
        InputDirEntry.configure(state = DISABLED)
        InputDirChooseBtn.configure(state = DISABLED)
        DirCalcBtn.configure(state = DISABLED)
        FilesCountLbl.configure(state = DISABLED)
        
        SaveNameLbl.configure(state = DISABLED)
        SaveNameEntry.configure(state = DISABLED)
        SaveDirBtn.configure(state = DISABLED)
        SaveDirEntry.configure(state = DISABLED)

        ModeSaveEnum.configure(state = DISABLED)
        ModeMoveCompleted.configure(state = DISABLED)
        
        ModeMergeAll.configure(state = DISABLED)
        ModeMergeByC.configure(state = DISABLED)
        ModeMergeByE.configure(state = DISABLED)
        
        MergeDivisorLbl.configure(state = DISABLED)
        MergeDivisorSpin.configure(state = DISABLED)
        MergeQuotientLbl.configure(state = DISABLED)
        MergeRemainderLbl.configure(state = DISABLED)
        
        MergeBtn.configure(state = DISABLED)
    else:
        InputDirPathLbl.configure(state = NORMAL)
        InputDirEntry.configure(state = NORMAL)
        InputDirChooseBtn.configure(state = NORMAL)
        DirCalcBtn.configure(state = NORMAL)
        FilesCountLbl.configure(state = NORMAL)
        
        SaveNameLbl.configure(state = NORMAL)
        SaveNameEntry.configure(state = NORMAL)
        SaveDirBtn.configure(state = NORMAL)
        SaveDirEntry.configure(state = NORMAL)
    
        ModeSaveEnum.configure(state = NORMAL)
        ModeMoveCompleted.configure(state = NORMAL)
        
        ModeMergeAll.configure(state = NORMAL)
        ModeMergeByC.configure(state = NORMAL)
        ModeMergeByE.configure(state = NORMAL)
        
        MergeDivisorLbl.configure(state = NORMAL)
        MergeDivisorSpin.configure(state = NORMAL)
        MergeQuotientLbl.configure(state = NORMAL)
        MergeRemainderLbl.configure(state = NORMAL)
        
        MergeBtn.configure(state = NORMAL)

def TimeUpdater():
    global StartedMergeTime

    while MergeInProgress:
        CreateDocTime = time.time()
        result = CreateDocTime - StartedMergeTime
        result = datetime2.timedelta(seconds=round(result))
        MergeTimeLbl.config(text = str(result))
        time.sleep(0.01)




def CountFiles():
    global inputdir
    global FilesArray
    global ValidFilesCount
    global MergeMode
    global MergeDivisor
    global MergeQuotient
    global MergeRemainder
    FilesArray = []
    MergeDivisor = MergeDivisorSpin.get()
    MergeDivisor = int(MergeDivisor)
    MergeQuotient = 0
    MergeRemainder = 0
    print('======')
    
    if CheckInputDir():
    
        inputdir = InputDirEntry.get()
        FilesArray.clear()
        for file in os.listdir(inputdir):
            if file.endswith(".pdf"):
                FilesArray.append(os.path.join(inputdir, file))
        ValidFilesCount = len(FilesArray)
        FilesCountLbl.config(text = 'Количество файлов PDF: ' + str(len(FilesArray)))
        print('CF: number of valid files -', str(ValidFilesCount))
        
        if MergeMode.get() == 1:
            ValidFilesCountDivb = int(ValidFilesCount)
            while ValidFilesCountDivb%MergeDivisor != 0:
                ValidFilesCountDivb = ValidFilesCountDivb - 1
                
            if MergeDivisor > ValidFilesCount:
                MergeDivisor = ValidFilesCount
                MergeDivisorSpin.delete(0, END)
                MergeDivisorSpin.insert(0, MergeDivisor)
                
            MergeQuotient = round(ValidFilesCountDivb / MergeDivisor)
            MergeRemainder = round(ValidFilesCount - (MergeQuotient * MergeDivisor))
            
            MergeDivisor, MergeQuotient = MergeQuotient, MergeDivisor
            
            if MergeRemainder > 0:
                MergeDivisor = MergeDivisor +1
                MergeQuotientLbl.config(text = ('Количество частей: ' + str(MergeDivisor)))
                MergeRemainderLbl.config(text = ('Документов в последней: ' + str(MergeRemainder)))
            else:
                MergeQuotientLbl.config(text = ('Количество частей: ' + str(MergeDivisor)))
                MergeRemainderLbl.config(text = ('Документов в последней: ' + str(MergeQuotient)))
                
            
            print('CF: MergeDivisor -', str(MergeDivisor))
            print('CF: MergeQuotient -', str(MergeQuotient))
            print('CF: MergeRemainder -', str(MergeRemainder))
            

        
        if MergeMode.get() == 2:
            ValidFilesCountDivb = int(ValidFilesCount)
            while ValidFilesCountDivb%MergeDivisor != 0:
                ValidFilesCountDivb = ValidFilesCountDivb - 1
                
            if MergeDivisor > ValidFilesCount:
                MergeDivisor = ValidFilesCount
                MergeDivisorSpin.delete(0, END)
                MergeDivisorSpin.insert(0, MergeDivisor)
                
            MergeQuotient = round(ValidFilesCountDivb / MergeDivisor)
            MergeRemainder = MergeQuotient + round(ValidFilesCount - (MergeQuotient * MergeDivisor))
            
            print('CF: MergeDivisor -', str(MergeDivisor))
            print('CF: MergeQuotient -', str(MergeQuotient))
            print('CF: MergeRemainder -', str(MergeRemainder))
            
            MergeQuotientLbl.config(text = ('Документов в одной: ' + str(MergeQuotient)))
            MergeRemainderLbl.config(text = ('Документов в последней: ' + str(MergeRemainder)))
    else:
        FilesCountLbl.config(text = 'Неправильный путь')




def StartMergingThread():
    global inputdir
    global CreateDocTime
    global savedir
    global savename
    global outputfile
    global InputPath
    global OutputPath
    global StartedMerge
    
    inputdir = InputDirEntry.get()
    savedir = SaveDirEntry.get()
    savename = SaveNameEntry.get()
    
    print('======')
    if inputdir:
        print("SMT: inputdir exists")
        isinputdir = os.path.isdir(inputdir)
        print('SMT: inputdir is valid -', isinputdir)
        if isinputdir:
            CountFiles()
            print('======')
            if ValidFilesCount > 0:
                print("SMT: valid files available")
                if savedir:
                    print("SMT: savedir exists")
                    issavedir = os.path.isdir(savedir)
                    print('SMT: inputdir is valid -', issavedir)
                    if issavedir:
                        issavedirexist = os.path.exists(savedir)
                        if not issavedirexist:
                            os.makedirs(savedir)
                        if savename:
                            print("SMT: savename exists")
                            
                            DateTimeUpdate(True)
                            InputPath = Path(inputdir)
                            
                            print('SMT: starting merge')
                            mergethread = Thread(target=PDFmerge)
                            mergethread.start()
                            timethread = Thread(target=TimeUpdater)
                            timethread.start()
                            mergethread = ""
                            timethread = ""
                        else:
                            print("SMT: savename is empty")
                            
                            savename = SavenameGenerate(True, SettAddTime, savename, 0)
                            SaveNameEntry.delete(0,END)
                            SaveNameEntry.insert(0,SavenameGenerate(True, SettAddTime, savename, 0))
                            
                            DateTimeUpdate(True)
                            InputPath = Path(inputdir)
                            
                            print('SMT: starting merge')
                            mergethread = Thread(target=PDFmerge)
                            mergethread.start()
                            timethread = Thread(target=TimeUpdater)
                            timethread.start()
                            mergethread = ""
                            timethread = ""
                else:
                    print("SMT: savedir is empty")
                    savedir = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'Merger') 
                    #savedir = r"\\zorb-srv\Operators\ORBScan\Merger"
                    #savedir = r"C:\Users\ORB User\Desktop\Merger"
                    issavedirexist = os.path.exists(savedir)
                    if not issavedirexist:
                        os.makedirs(savedir)
                    SaveDirEntry.delete(0,END)
                    SaveDirEntry.insert(0,savedir)
                    StartMergingThread()
                    
            else:
                print("SMT: no valid files")
    else:
        print("SMT: inputdir is empty")
        FilesCountLbl.config(text = 'Неправильный путь')




def PDFmerge():
    global MergeInProgress
    global StartedMergeTime
    global InputPath
    global OutputPath
    global savedir
    global savename

    global MergeMode

    global MergeDivisor
    global MergeQuotient
    global MergeRemainder

    global outputfilesenum
    global OutputPathFilesEnum
    
    global CreateFilesEnum
    global MoveCompletedFiles
    
    global MoveFilesPath
    global RevisedMoveFilesPath


    print('====== PDFM ======')
    MergeInProgress = True
    StartedMergeTime = time.time()
    BlockGUI(True)

    print('PDFM: CreateFilesEnum', CreateFilesEnum.get())
    print('PDFM: MoveCompletedFiles', MoveCompletedFiles.get())

    if MergeMode.get() == 0:
        print('PDFM: Merge all')
        
        outputfile = (SavenameGenerate(False, False, savename, 0) + ".pdf")
        OutputPath = Path(savedir, outputfile)
        RevisedOutputPath = OutputPath.as_posix()
        print('PDFM: InputPath -', InputPath)
        print('PDFM: OutputPath -', OutputPath)
        print('PDFM: RevisedOutputPath -', RevisedOutputPath)
        
        # Adding files to merge task
        MergeProgressLbl.config(text = '')
        #pdfmerger = PdfFileMerger()
        for i in range(0, len(FilesArray)):
            #pdfmerger.append(FilesArray[i])
            
            MergeStatusLbl.config(text = ('Отправка на печать: '+str(i+1)))
            print('========================================')
            print('Отправка на печать - i: '+str(i))
            print('Файл: ' +str(FilesArray[i]))
            win32api.ShellExecute(0, "print", FilesArray[i], None,  ".",  0)
            
        # Starting merge task
        MergeStatusLbl.config(text = 'Данные отправлены !')
        #pdfmerger.write(RevisedOutputPath)
        #pdfmerger = ""
        if CreateFilesEnum.get() == 1:
            filesenum.close()
        if MoveCompletedFiles.get() == 1:
            print('========================================')
            MergeStatusLbl.config(text = 'Перемещение файлов...')
            print('Перемещение файлов...')
            MoveFilesPath = Path(savedir, savename)
            RevisedMoveFilesPath = MoveFilesPath.as_posix()
            ismovedirexist = os.path.exists(RevisedMoveFilesPath)
            if not ismovedirexist:
                os.makedirs(RevisedMoveFilesPath)
            for i in range(0, len(FilesArray)):
                #print('Начальный файл: ' +str(FilesArray[i]))
                filename = Path(FilesArray[i])
                MoveFilesPath = Path(savedir, savename, filename.name)
                RevisedMoveFilesPath = MoveFilesPath.as_posix()
                #print('Конечный файл: ' +str(RevisedMoveFilesPath))
                #shutil.move(FilesArray[i], RevisedMoveFilesPath)
                shutil.copyfile(FilesArray[i], RevisedMoveFilesPath)
        MergeStatusLbl.config(text = 'Обьединение завершено !')
        

        




    if MergeMode.get() >= 1:
        # Main loop part
        print('PDFM: Merge by parts')
        print('MergeDivisor: '+str(MergeDivisor))
        print('Loop part:')

        BlockStart = 1
        BlockEnd = 1
        MergeProgressLbl.config(text = ('0'+'/'+str(MergeDivisor)))
        
           
        print('=================== Zero K =======================')
        for k in range(0, MergeDivisor):
            pdfmerger = PdfFileMerger()
            BlockStart = BlockEnd
            BlockEnd = MergeQuotient*k
            
            print('Part k: '+str(k))
            print('BlockStart: '+str(BlockStart))
            print('BlockEnd: '+str(BlockEnd))
            i = BlockStart + 1
            l =  i
            
            if k > 0:
            
                if CreateFilesEnum.get() == 1:
                    outputfilesenum = (SavenameGenerate(False, False, savename, k) + " files list" + ".txt")
                    OutputPathFilesEnum = Path(savedir, outputfilesenum)
                    RevisedOutputPathFilesEnum = OutputPathFilesEnum.as_posix()
                    try:
                        os.remove(RevisedOutputPathFilesEnum)
                        print("PDFM: Existing enum file removed")
                    except IOError:
                        print("PDFM: No enum file!")
                    filesenum = open(RevisedOutputPathFilesEnum, 'a')
            
                while i <= BlockEnd:
                    pdfmerger.append(FilesArray[i-1])
                    if CreateFilesEnum.get() == 1:
                        filesenum.write(FilesArray[i-1])
                        filesenum.write('\n')
                    MergeStatusLbl.config(text = ('Добавление в задачу: '+str(i+1)))
                    print('======================')
                    print('Добавление в задачу - i: '+str(i))
                    print('Добавление в задачу - pos: '+str(i-1))
                    print('Файл: ' +str(FilesArray[i-1]))
                    i += 1
                
                MergeStatusLbl.config(text = ('Обьединение документа № ' +str(k)))
                outputfile = (SavenameGenerate(False, False, savename, k) + ".pdf")
                OutputPath = Path(savedir, outputfile)
                RevisedOutputPath = OutputPath.as_posix()
                pdfmerger.write(RevisedOutputPath)
                pdfmerger = ""
                if CreateFilesEnum.get() == 1:
                    filesenum.close()
                print('Итоговый файл: ' +str(RevisedOutputPath))
                
                if MoveCompletedFiles.get() == 1:
                    print('========================================')
                    print('Создание папки под файлы...')
                    MoveFilesPath = Path(savedir, SavenameGenerate(False, False, savename, k))
                    RevisedMoveFilesPath = MoveFilesPath.as_posix()
                    ismovedirexist = os.path.exists(RevisedMoveFilesPath)
                    if not ismovedirexist:
                        os.makedirs(RevisedMoveFilesPath)
                    MergeStatusLbl.config(text = 'Перемещение файлов...')
                    print('Перемещение файлов...')
                    while l <= BlockEnd:
                        #print('Начальный файл: ' +str(FilesArray[l-1]))
                        filename = Path(FilesArray[l-1])
                        MoveFilesPath = Path(savedir, SavenameGenerate(False, False, savename, k), filename.name)
                        RevisedMoveFilesPath = MoveFilesPath.as_posix()
                        #print('Конечный файл: ' +str(RevisedMoveFilesPath))
                        #shutil.move(FilesArray[l-1], RevisedMoveFilesPath)
                        shutil.copyfile(FilesArray[l-1], RevisedMoveFilesPath)
                        l += 1
                        
                MergeProgressLbl.config(text = (str(k)+'/'+str(MergeDivisor)))
                print('=================== Next K =======================')
        
        # Main remainder part
        print('Remainder part:')
        BlockStart = BlockEnd+1
        BlockEnd = len(FilesArray)+1
        pdfmerger = PdfFileMerger()
        print('BlockStart: '+str(BlockStart))
        print('BlockEnd: '+str(BlockEnd))
        
        if CreateFilesEnum.get() == 1:
            outputfilesenum = (SavenameGenerate(False, False, savename, k+1) + " files list" + ".txt")
            OutputPathFilesEnum = Path(savedir, outputfilesenum)
            RevisedOutputPathFilesEnum = OutputPathFilesEnum.as_posix()
            try:
                os.remove(RevisedOutputPathFilesEnum)
                print("PDFM: Existing enum file removed")
            except IOError:
                print("PDFM: No enum file!")
            filesenum = open(RevisedOutputPathFilesEnum, 'a')
        
        for i in range(BlockStart, BlockEnd):
            pdfmerger.append(FilesArray[i-1])
            if CreateFilesEnum.get() == 1:
                filesenum.write(FilesArray[i-1])
                filesenum.write('\n')
            MergeStatusLbl.config(text = ('Добавление в задачу: '+str(i+1)))
            print('======================')
            print('Добавление в задачу - i: '+str(i))
            print('Добавление в задачу - pos: '+str(i-1))
            print('Файл: ' +str(FilesArray[i-1]))
            
        MergeStatusLbl.config(text = 'Обьединение последнего документа...') 
        outputfile = (SavenameGenerate(False, False, savename, MergeDivisor) + ".pdf")
        OutputPath = Path(savedir, outputfile)
        RevisedOutputPath = OutputPath.as_posix()
        pdfmerger.write(RevisedOutputPath)
        pdfmerger = ""
        if CreateFilesEnum.get() == 1:
            filesenum.close()
        print('Последний итоговый файл: ' +str(RevisedOutputPath))
            
        if MoveCompletedFiles.get() == 1:
            print('========================================')
            print('Создание папки под файлы...')
            MoveFilesPath = Path(savedir, SavenameGenerate(False, False, savename, MergeDivisor))
            RevisedMoveFilesPath = MoveFilesPath.as_posix()
            ismovedirexist = os.path.exists(RevisedMoveFilesPath)
            if not ismovedirexist:
                os.makedirs(RevisedMoveFilesPath)
            MergeStatusLbl.config(text = 'Перемещение файлов...')
            print('Перемещение файлов...')
            for l in range(BlockStart, BlockEnd):
                #print('Начальный файл: ' +str(FilesArray[l-1]))
                filename = Path(FilesArray[l-1])
                MoveFilesPath = Path(savedir, SavenameGenerate(False, False, savename, MergeDivisor), filename.name)
                RevisedMoveFilesPath = MoveFilesPath.as_posix()
                #print('Конечный файл: ' +str(RevisedMoveFilesPath))
                #shutil.move(FilesArray[l-1], RevisedMoveFilesPath)
                shutil.copyfile(FilesArray[l-1], RevisedMoveFilesPath)
        MergeStatusLbl.config(text = 'Выполнение завершено !')
        
        MergeProgressLbl.config(text = (str(MergeDivisor)+'/'+str(MergeDivisor)))

    print('========== PDFM: tasks finished ! ============')
    MergeInProgress = False
    BlockGUI(False)






if __name__ == '__main__':
    main()
