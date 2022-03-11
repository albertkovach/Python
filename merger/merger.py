from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from PyPDF2 import PdfFileMerger
from threading import Thread
from datetime import datetime
from pathlib import Path
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




def main():
    global root
    root = Tk()
    root.resizable(False, False)
        
    scrnw = root.winfo_screenwidth()
    scrnh = root.winfo_screenheight()
    scrnw = scrnw//2
    scrnh = scrnh//2
    scrnw = scrnw - 175
    scrnh = scrnh - 100
    root.geometry('350x200+{}+{}'.format(scrnw, scrnh))
        
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
        fx = 20
        fy = 40

        global InputDirPathLbl
        InputDirPathLbl = Label(text="Выберите папку с файлами:", background="white")
        InputDirPathLbl.place(x=fx -4, y=fy -20)
        
        global InputDirEntry
        InputDirEntry = Entry(fg="black", bg="white", width=40)
        InputDirEntry.place(x=fx, y=fy)
        
        global InputDirChooseBtn
        InputDirChooseBtn = Button(text='Выбор', command=InputDirChoose)
        InputDirChooseBtn.place(x=fx +250, y=fy -1, height=20)

        global DirCalcBtn
        DirCalcBtn = Button(text='Посчитать', command=CountFiles)
        DirCalcBtn.place(x=fx, y=fy +22, height=20)

        global FilesCountLbl
        FilesCountLbl = Label(text="", background="white")
        FilesCountLbl.place(x=fx +75, y=fy +22)


        global SaveNameLbl
        SaveNameLbl = Label(text="Введите имя итогового файла:", background="white")
        SaveNameLbl.place(x=fx -4, y=fy +52)

        global SaveNameEntry
        SaveNameEntry = Entry(fg="black", bg="white", width=20)
        SaveNameEntry.place(x=fx +176, y=fy +53)

        global SaveDirBtn
        SaveDirBtn = Button(text="Выбрать путь", command=SaveDirChoose)
        SaveDirBtn.place(x=fx, y=fy +75, height=20)

        global SaveDirEntry
        SaveDirEntry = Entry(fg="black", bg="white", width=34)
        SaveDirEntry.place(x=fx +93, y=fy +76)


        global MergeBtn
        MergeBtn = Button(text="Выполнить слияние", command=StartMergingThread)
        MergeBtn.place(x=fx, y=fy +120, height=20)




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
                        else:
                            print("SMT: savename is empty")
                            now = datetime.now()
                            savename = ("Merged_"+str(now.date())+"_"+str(now.hour)+"-"+str(now.minute)+"-"+str(now.second))
                            SaveNameEntry.delete(0,END)
                            SaveNameEntry.insert(0,savename)
                            StartMergingThread()
                else:
                    print("SMT: savedir is empty")
                    #desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
                    #savedir = r"\\zorb-srv\Operators\ORBScan\merged-pdf"
                    savedir = r"C:\Users\ORB User\Desktop"
                    SaveDirEntry.delete(0,END)
                    SaveDirEntry.insert(0,savedir)
                    StartMergingThread()
                    
            else:
                print("SMT: no valid files")
    else:
        print("SMT: inputdir is empty")
        FilesCountLbl.config(text = 'Неправильный путь')




def PDFmerge():
    print('====== PDFM ======')
    global InputPath
    global OutputPath
    print('PDFM: InputPath -', InputPath)
    print('PDFM: OutputPath -', OutputPath)
    
    RevisedOutputPath = OutputPath.as_posix()
    print('PDFM: RevisedOutputPath -', RevisedOutputPath)

    pdfmerger = PdfFileMerger()

    if len(FilesArray) > 0 :
        for i in range(0, len(FilesArray)):
            pdfmerger.append(FilesArray[i])
            print(i)
            print(FilesArray[i])
        pdfmerger.write(RevisedOutputPath)



if __name__ == '__main__':
    main()

