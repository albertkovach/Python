from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from PyPDF2 import PdfFileMerger
from threading import Thread
from datetime import datetime
from pathlib import Path
import datetime as datetime2
import time
import os
import re


global root
global inputdir
global savedir
global savename
global outputfile

global ValidFilesCount
global FilesArray

global InputPath
global OutputPath
global mergethread
global timethread

global MergeInProgress
global StartedMergeTime
MergeInProgress = False



def main():
    global root
    root = Tk()
    root.resizable(False, False)
        
    scrnw = root.winfo_screenwidth()
    scrnh = root.winfo_screenheight()
    scrnw = scrnw//2
    scrnh = scrnh//2
    scrnw = scrnw - 175
    scrnh = scrnh - 125
    root.geometry('350x250+{}+{}'.format(scrnw, scrnh))
        
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
        fx = 20
        fy = 35

        global InputDirPathLbl
        InputDirPathLbl = Label(text="Выберите папку с файлами:", background="white")
        InputDirPathLbl.place(x=16, y=10)
        
        global InputDirEntry
        InputDirEntry = Entry(fg="black", bg="white", width=40)
        InputDirEntry.place(x=20, y=30)
        
        global InputDirChooseBtn
        InputDirChooseBtn = Button(text='Выбор', command=InputDirChoose)
        InputDirChooseBtn.place(x=270, y=29, height=20)

        global DirCalcBtn
        DirCalcBtn = Button(text='Посчитать', command=CountFiles)
        DirCalcBtn.place(x=20, y=52, height=20)

        global FilesCountLbl
        FilesCountLbl = Label(text="", background="white")
        FilesCountLbl.place(x=95, y=52)


        global SaveNameLbl
        SaveNameLbl = Label(text="Введите имя итогового файла:", background="white")
        SaveNameLbl.place(x=16, y=87)

        global SaveNameEntry
        SaveNameEntry = Entry(fg="black", bg="white", width=20)
        SaveNameEntry.place(x=196, y=88)

        global SaveDirBtn
        SaveDirBtn = Button(text="Выбрать путь", command=SaveDirChoose)
        SaveDirBtn.place(x=20, y=110, height=20)

        global SaveDirEntry
        SaveDirEntry = Entry(fg="black", bg="white", width=34)
        SaveDirEntry.place(x=113, y=111)


        global MergeBtn
        MergeBtn = Button(text="Выполнить склейку", command=StartMergingThread)
        MergeBtn.place(x=20, y=175, height=20)

        global MergeStatusLbl
        MergeStatusLbl = Label(text="", background="white")
        MergeStatusLbl.place(x=145, y=175)
        
        global MergeTimeLbl
        MergeTimeLbl = Label(text="", background="white")
        MergeTimeLbl.place(x=145, y=192)





def InputDirChoose(): 
    global inputdir
    
    inputdir = filedialog.askdirectory(title="Выбрать папку")
    print('======')
    if inputdir:
        InputDirEntry.delete(0,END)
        InputDirEntry.insert(0,inputdir)
        print('DChs: inputdir -', inputdir)
        CountFiles()
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




def CountFiles():
    global inputdir
    global FilesArray
    global ValidFilesCount
    
    FilesArray = []
    isinputdir = os.path.isdir(inputdir)
    print('======')
    print('CF: inputdir is valid - ', isinputdir)
    if isinputdir:
        FilesArray.clear()
        for file in os.listdir(inputdir):
            if file.endswith(".pdf"):
                FilesArray.append(os.path.join(inputdir, file))
        ValidFilesCount = len(FilesArray)
        FilesCountLbl.config(text = 'Количество файлов PDF: ' + str(len(FilesArray)))
        print('CF: number of valid files -', str(ValidFilesCount))
    else:
        FilesCountLbl.config(text = 'Неправильный путь')




def StartMergingThread():
    global inputdir
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
                        if savename:
                            print("SMT: savename exists")
                            InputPath = Path(inputdir)
                            outputfile = (str(savename) + ".pdf")
                            OutputPath = Path(savedir, outputfile)
                            print('SMT: starting merge')
                            mergethread = Thread(target=PDFmerge)
                            mergethread.start()
                            timethread = Thread(target=TimeUpdater)
                            timethread.start()
                            mergethread = ""
                            timethread = ""
                        else:
                            print("SMT: savename is empty")
                            now = datetime.now()
                            savename = ("Merged PDF "+str(now.day)+"."+str(now.month)+"."+str(now.year)+" "+str(now.hour)+"-"+str(now.minute)+"-"+str(now.second))
                            SaveNameEntry.delete(0,END)
                            SaveNameEntry.insert(0,savename)
                            StartMergingThread()
                else:
                    print("SMT: savedir is empty")
                    #desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
                    #savedir = r"\\zorb-srv\Operators\ORBScan\merged-pdf"
                    savedir = r"C:\Users\ORB User\Desktop\1"
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

    print('====== PDFM ======')
    MergeInProgress = True
    StartedMergeTime = time.time()
    print('PDFM: InputPath -', InputPath)
    print('PDFM: OutputPath -', OutputPath)
    
    RevisedOutputPath = OutputPath.as_posix()
    print('PDFM: RevisedOutputPath -', RevisedOutputPath)

    pdfmerger = PdfFileMerger()

    if len(FilesArray) > 0 :
        for i in range(0, len(FilesArray)):
            pdfmerger.append(FilesArray[i])
            MergeStatusLbl.config(text = ('Добавление в задачу: '+str(i+1)))
            print(i)
            print(FilesArray[i])
            
        MergeStatusLbl.config(text = 'Обьединение документа...') 
        pdfmerger.write(RevisedOutputPath)
        MergeStatusLbl.config(text = 'Обьединение завершено !') 
        
    MergeInProgress = False



def TimeUpdater():
    global StartedMergeTime

    while MergeInProgress:
        CurrentTime = time.time()
        result = CurrentTime - StartedMergeTime
        result = datetime2.timedelta(seconds=round(result))
        MergeTimeLbl.config(text = str(result))
        time.sleep(0.01)


if __name__ == '__main__':
    main()
