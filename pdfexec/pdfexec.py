from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk

#сколько не использовано кодов/файлов в startmoving
#перегрузка данных на кнопку

#
#

from pathlib import Path
import shutil


import os
global root

global MoveInputDir
global MoveInputDirTxt
global MoveAvlFolders
global MoveAvlFullFolders
global MoveAvlFolderRestFiles

global MoveSaveDir
global MoveSaveDirTxt
global MoveBarcodeFile
global MoveBarcodeTxt

global MoveIsInputSel
global MoveIsOutputSel
global MoveIsBarcodeSel

global MoveFilesArray
global MoveValidFilesCount
global MoveFileAmounToMove
MoveFileAmounToMove = 10

global MoveBarcodeArray
global MoveBarcodeCount



global CurrentUIpage





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
        self.parent.title("")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        global MoveIsInputSel
        global MoveIsOutputSel
        global MoveIsBarcodeSel
        MoveIsInputSel = False
        MoveIsOutputSel = False
        MoveIsBarcodeSel = False
        
        s = ttk.Style()
        s.configure('.', background="white")
        


        # >>>>>>>>>>>>>>>> Mode header <<<<<<<<<<<<<<<<
        # =============================================
        global MainModeLbl
        MainModeLbl = Label(text="Разделение по коробам", background="white", font=("Arial", 12))
        MainModeLbl.place(x=16, y=2)
        
        global MainModeBackBtn
        MainModeBackBtn = Button(text='Назад', command=SetModePageBack)
        
        global MainModeSeparator
        MainModeSeparator = ttk.Separator(root, orient='horizontal')
        MainModeSeparator.place(relx=0, y=32, relwidth=1, relheight=1)


        # >>>>>>>>>>>>>>> Mode selector <<<<<<<<<<<<<<<
        # =============================================
        global SetModeDividePdfBtn
        SetModeDividePdfBtn = Button(text='Разделить на счет-фактуры', command=SetModeDividePdf, font=("Arial", 11))
        
        global SetModeMoveByBoxBtn
        SetModeMoveByBoxBtn = Button(text='Раскидать по коробам', command=SetModeMoveByBox, font=("Arial", 11))
        
        global SetModeSortAndMergeBtn
        SetModeSortAndMergeBtn = Button(text='Объединение в коробах', command=SetModeSortAndMerge, font=("Arial", 11))








        # >>>> ModeMove input folder widgets block <<<<
        # =============================================
        global MoveInputDirPathLbl
        MoveInputDirPathLbl = Label(text="Выберите папку с файлами:", background="white", font=("Arial", 10))
        
        global MoveInputDirTxt
        MoveInputDirTxt = StringVar()
        global MoveInputDirEntry
        MoveInputDirEntry = Entry(fg="black", bg="white", width=46, textvariable=MoveInputDirTxt)
        MoveInputDirEntry.configure(state = DISABLED)
        
        global MoveInputDirChooseBtn
        MoveInputDirChooseBtn = Button(text='Выбор', command=MoveChooseInputDir)

        global MoveFilesCountLbl
        MoveFilesCountLbl = Label(text="", background="white")
        
        
        # >>>> ModeMove output folder widgets block <<<
        # =============================================
        global MoveSaveDirBtnLbl
        MoveSaveDirBtnLbl = Label(text="Папка для коробов:", background="white", font=("Arial", 10))
        
        global MoveSaveDirBtn
        MoveSaveDirBtn = Button(text="Выбор", command=MoveChooseOutputDir)

        global MoveSaveDirTxt
        MoveSaveDirTxt = StringVar()
        global MoveSaveDirEntry
        MoveSaveDirEntry = Entry(fg="black", bg="white", width=46, textvariable=MoveSaveDirTxt)
        MoveSaveDirEntry.configure(state = DISABLED)
        
        
        # >>>>>>> ModeMove barcode widgets block <<<<<<
        # =============================================
        global MoveBarcodeSelLbl
        MoveBarcodeSelLbl = Label(text="Файл с штрихкодами:", background="white", font=("Arial", 10))
        
        global MoveBarcodeSelBtn
        MoveBarcodeSelBtn = Button(text="Выбор", command=MoveChooseBarcodeFile)

        global MoveBarcodeTxt
        MoveBarcodeTxt = StringVar()
        global MoveBarcodeSelEntry
        MoveBarcodeSelEntry = Entry(fg="black", bg="white", width=46, textvariable=MoveBarcodeTxt)
        MoveBarcodeSelEntry.configure(state = DISABLED)
        
        global MoveBarcodeCountLbl
        MoveBarcodeCountLbl = Label(text="", background="white")
        
        global MoveRunDivisionBtn
        MoveRunDivisionBtn = Button(text="Выполнить", command=MoveStartMoving)
        MoveRunDivisionBtn.configure(state = DISABLED)
        
        
        
        
        
        
        
        
        
        global CurrentUIpage
        CurrentUIpage = 2
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
    global CurrentUIpage
    CurrentUIpage = 0
    UIswitcher()

def UIswitcher():
    global CurrentUIpage
    if CurrentUIpage == 0:
        MainModeLbl.config(text = 'Выберите режим:')
        MainModeBackBtn.place_forget()
        
        SetModeDividePdfBtn.place(x=55, y=60, width = 300, height=27)
        SetModeMoveByBoxBtn.place(x=55, y=110, width = 300, height=27)
        SetModeSortAndMergeBtn.place(x=55, y=160, width = 300, height=27)

        MoveInputDirPathLbl.place_forget()
        MoveInputDirEntry.place_forget()
        MoveInputDirChooseBtn.place_forget()
        MoveFilesCountLbl.place_forget()
        MoveSaveDirBtnLbl.place_forget()
        MoveSaveDirBtn.place_forget()
        MoveSaveDirEntry.place_forget()
        MoveBarcodeSelLbl.place_forget()
        MoveBarcodeSelBtn.place_forget()
        MoveBarcodeSelEntry.place_forget()
        MoveBarcodeCountLbl.place_forget()
        MoveRunDivisionBtn.place_forget()
        
    if CurrentUIpage == 1:
        MainModeLbl.config(text = 'Разделение по счет-фактурам:')
        MainModeBackBtn.place(x=310, y=6, height=20)
        
        SetModeDividePdfBtn.place_forget()
        SetModeMoveByBoxBtn.place_forget()
        SetModeSortAndMergeBtn.place_forget()

        MoveInputDirPathLbl.place_forget()
        MoveInputDirEntry.place_forget()
        MoveInputDirChooseBtn.place_forget()
        MoveFilesCountLbl.place_forget()
        MoveSaveDirBtnLbl.place_forget()
        MoveSaveDirBtn.place_forget()
        MoveSaveDirEntry.place_forget()
        MoveBarcodeSelLbl.place_forget()
        MoveBarcodeSelBtn.place_forget()
        MoveBarcodeSelEntry.place_forget()
        MoveBarcodeCountLbl.place_forget()
        MoveRunDivisionBtn.place_forget()
        
    if CurrentUIpage == 2:
        MainModeLbl.config(text = 'Разделение по коробам:')
        MainModeBackBtn.place(x=310, y=6, height=20)
        
        SetModeDividePdfBtn.place_forget()
        SetModeMoveByBoxBtn.place_forget()
        SetModeSortAndMergeBtn.place_forget()

        MoveInputDirPathLbl.place(x=16, y=7+40)
        MoveInputDirEntry.place(x=20, y=30+40)
        MoveInputDirChooseBtn.place(x=305, y=29+40, height=20)
        MoveFilesCountLbl.place(x=17, y=49+40)
        MoveSaveDirBtnLbl.place(x=16, y=87+40)
        MoveSaveDirBtn.place(x=20, y=110+40, height=20)
        MoveSaveDirEntry.place(x=75, y=111+40)
        MoveBarcodeSelLbl.place(x=16, y=141+40)
        MoveBarcodeSelBtn.place(x=20, y=164+40, height=20)
        MoveBarcodeSelEntry.place(x=75, y=165+40)
        MoveBarcodeCountLbl.place(x=17, y=185+40)
        MoveRunDivisionBtn.place(x=20, y=225+40, width=90, height=25)
        
    if CurrentUIpage == 3:
        MainModeLbl.config(text = 'Объединение счет-фактур:')
        MainModeBackBtn.place(x=310, y=6, height=20)
        
        SetModeDividePdfBtn.place_forget()
        SetModeMoveByBoxBtn.place_forget()
        SetModeSortAndMergeBtn.place_forget()

        MoveInputDirPathLbl.place_forget()
        MoveInputDirEntry.place_forget()
        MoveInputDirChooseBtn.place_forget()
        MoveFilesCountLbl.place_forget()
        MoveSaveDirBtnLbl.place_forget()
        MoveSaveDirBtn.place_forget()
        MoveSaveDirEntry.place_forget()
        MoveBarcodeSelLbl.place_forget()
        MoveBarcodeSelBtn.place_forget()
        MoveBarcodeSelEntry.place_forget()
        MoveBarcodeCountLbl.place_forget()
        MoveRunDivisionBtn.place_forget()





def MoveChooseInputDir():
    global MoveInputDir
    global MoveFileAmounToMove
    global MoveIsInputSel
    
    global MoveFilesArray
    global MoveValidFilesCount
    
    global MoveAvlFolders
    global MoveAvlFullFolders
    global MoveAvlFolderRestFiles
    
    MoveInputDir = filedialog.askdirectory()
    
    if MoveInputDir:
        MoveInputDirTxt.set(str(MoveInputDir))
        print('MoveInputDir :', MoveInputDir)
        
        MoveFilesArray = []
        ScanFolder(MoveInputDir, ".pdf", MoveFilesArray)
        MoveValidFilesCount = len(MoveFilesArray)

        if MoveValidFilesCount > 0:
        
            MoveAvlFolders = 0
            MoveAvlFullFolders = len(MoveFilesArray)/MoveFileAmounToMove
            print('MoveAvlFullFolders:', str(MoveAvlFullFolders))
            
            if MoveAvlFullFolders - int(MoveAvlFullFolders) > 0:
                print('MoveAvlFolderRestFiles: ', str(MoveAvlFullFolders - int(MoveAvlFullFolders)))
                MoveAvlFolderRestFiles = True
                MoveAvlFolders = int(MoveAvlFullFolders + 1)
                lbltext = ("Файлов PDF: {0}, нужно {1} кодов, последний - неполный".format(len(MoveFilesArray),MoveAvlFolders))
            else:
                MoveAvlFolders = int(MoveAvlFullFolders)
                lbltext = ("Файлов PDF: {0}, нужно {1} кодов".format(len(MoveFilesArray),MoveAvlFolders))
            MoveAvlFullFolders = int(MoveAvlFullFolders)
        
            MoveFilesCountLbl.config(text = lbltext)
            print('Number of valid files -', str(MoveValidFilesCount))
            MoveIsInputSel = True
            MoveCheckIfReady()
        else:
            MoveFilesCountLbl.config(text = 'Нет файлов PDF !')
            MoveIsInputSel = False
            MoveCheckIfReady()
    else:
        print('MoveInputDir not selected')
        


def MoveChooseOutputDir():
    global MoveSaveDir
    global MoveSaveDirTxt
    global MoveIsOutputSel
    MoveSaveDir = filedialog.askdirectory()

    if MoveSaveDir:
        MoveSaveDirTxt.set(str(MoveSaveDir))
        print('MoveSaveDir :', MoveSaveDir)
        
        MoveIsOutputSel = True
        MoveCheckIfReady()
    else:
        print('MoveSaveDir not selected')



def MoveChooseBarcodeFile():
    global MoveBarcodeFile
    global MoveBarcodeTxt
    global MoveBarcodeArray
    global MoveIsBarcodeSel
    
    MoveBarcodeFile = filedialog.askopenfilename(filetypes=(('text files', 'txt'),))
    
    if MoveBarcodeFile:
        MoveBarcodeTxt.set(str(MoveBarcodeFile))
        print('MoveBarcodeTxt :', MoveBarcodeFile)
        
        file = open(MoveBarcodeFile,'r')
        MoveBarcodeArray = []
        MoveBarcodeArray.clear()
        for line in file:
            MoveBarcodeArray.append(line.strip()) # We don't want newlines in our list
        
        if len(MoveBarcodeArray) > 0:
        
            result = CheckListDuplicates(MoveBarcodeArray)
            
            if result:
                MoveBarcodeCountLbl.config(text = 'Есть повторяющиеся коды !')
                print('MoveBarcodeFile contain duplicates !')
                MoveIsBarcodeSel = False
                MoveCheckIfReady()
            else:
                MoveBarcodeCountLbl.config(text = 'Кодов в файле: ' + str(len(MoveBarcodeArray)))
                for i in range(len(MoveBarcodeArray)):
                    print('  - bar:', str(MoveBarcodeArray[i]))

                MoveIsBarcodeSel = True
                MoveCheckIfReady()
        else:
            MoveBarcodeCountLbl.config(text = 'Штрихкодов не найдено!')
            print('MoveBarcodeFile is empty !')
            MoveIsBarcodeSel = False
            MoveCheckIfReady()
    else:
        print('MoveBarcodeFile not selected')

    
    
    
    
   
def MoveStartMoving():
    global MoveSaveDir
    global MoveBarcodeArray
    global MoveFileAmounToMove
    global MoveBarcodeFile
    
    global MoveFilesArray
    global MoveValidFilesCount
    
    global MoveIsInputSel
    global MoveIsOutputSel
    global MoveIsBarcodeSel
    
    global MoveAvlFolders
    folderstocreate = 0
    #global MoveAvlFullFolders
    #global MoveAvlFolderRestFiles
    
    
    print('StartMoving')
    
    if MoveAvlFolders < len(MoveBarcodeArray):
        folderstocreate = MoveAvlFolders
        restcodes = len(MoveBarcodeArray) - folderstocreate
    else:
        folderstocreate = len(MoveBarcodeArray)
    
    
    for i in range(folderstocreate):
        MoveOutputPath = Path(MoveSaveDir, MoveBarcodeArray[i])
        ismovedirexist = os.path.exists(MoveOutputPath)
        print('==== New folder:', str(MoveBarcodeArray[i]))
        
        if not ismovedirexist:
            os.makedirs(MoveOutputPath)

        temparray = []
        ScanFolder(MoveInputDir, ".pdf", temparray)
        if MoveFileAmounToMove < len(temparray):
            filestomove = MoveFileAmounToMove
        else:
            filestomove = len(temparray)
            
        for k in range(filestomove):
            filename = Path(MoveFilesArray[k])
            MoveFilesPath = Path(MoveSaveDir, MoveBarcodeArray[i], filename.name)
            RevisedMoveFilesPath = MoveFilesPath.as_posix()
            print("Folder: {0}, Num {1}, file: {2}".format(MoveBarcodeArray[i],k,str(filename.name)))
            shutil.move(MoveFilesArray[k], RevisedMoveFilesPath)
            #shutil.copyfile(MoveFilesArray[k], RevisedMoveFilesPath)
    
        MoveFilesArray = []
        MoveFilesArray.clear()
        k = 0
    
        ScanFolder(MoveInputDir, ".pdf", MoveFilesArray)
        MoveValidFilesCount = len(MoveFilesArray)
        MoveFilesCountLbl.config(text = 'Количество файлов PDF: ' + str(len(MoveFilesArray)))
        
    open(MoveBarcodeFile, 'w').close()
        
    messagebox.showinfo("", "Перемещение выполнено !")
    MoveIsBarcodeSel = False
    MoveBarcodeArray.clear
    
    MoveBarcodeSelEntry.configure(state = NORMAL)
    MoveBarcodeSelEntry.delete(0,END)
    MoveBarcodeSelEntry.configure(state = DISABLED)
    
    MoveBarcodeCountLbl.config(text = '')
    MoveRunDivisionBtn.configure(state = DISABLED)




def MoveCheckIfReady():
    global MoveSaveDir
    global MoveBarcodeArray
    global MoveAvlFolders
    
    global MoveIsInputSel
    global MoveIsOutputSel
    global MoveIsBarcodeSel
    
    if MoveIsInputSel and MoveIsOutputSel and MoveIsBarcodeSel:
    
        # Checking existing folders from savedir and barcodearray
        allfilesarray = []
        directoriesarray = []
        temparray = []
        ScanFolder(MoveSaveDir, "", allfilesarray)
        
        for i in range(len(allfilesarray)):
            if os.path.isdir(allfilesarray[i]):
                directoriesarray.append(str(Path(allfilesarray[i]).as_posix()))

        if len(directoriesarray) > 0:
            print('==== Present folders:', str(len(directoriesarray)))
            for i in range(len(directoriesarray)):
                ScanFolder(directoriesarray[i], ".pdf", temparray)
                print(" - PDFs: {0}, Folder {1}".format(len(temparray),directoriesarray[i]))
     
            temparray.clear
            temparray = MoveBarcodeArray
            print('==== Target folders:', str(len(temparray)))
            for i in range(len(temparray)):
                temparray[i] = Path(MoveSaveDir, temparray[i]).as_posix()
                print(' = Folder', str(temparray[i]))

            intersect = bool(set(temparray) & set(directoriesarray))
            print('Folders intersect:', str(intersect))
        
            if intersect:
                messagebox.showerror("", "Некоторые короба из списка штрихкодов уже созданы ! Переместите их или загрузите другой список")
                MoveIsBarcodeSel = False
                MoveBarcodeArray.clear
                
                MoveBarcodeSelEntry.configure(state = NORMAL)
                MoveBarcodeSelEntry.delete(0,END)
                MoveBarcodeSelEntry.configure(state = DISABLED)
                
                MoveBarcodeCountLbl.config(text = '')
                MoveRunDivisionBtn.configure(state = DISABLED)
            else:
                # Folders exists in savedir, but not from barcodesarray
                MoveRunDivisionBtn.configure(state = NORMAL)
                
        # All good, checking how many barcodes remains after work
        else:
            restcodes = len(MoveBarcodeArray) - folderstocreate
            if restcodes > 0:
                countlbl = ("Кодов в файле: {0}, останется {1} кодов".format(len(MoveBarcodeArray),restcodes))
            else:
                #filesremain = len(MoveFilesArray) - (
                countlbl = ("Кодов в файле: {0}, останется файлов: {1}".format(len(MoveBarcodeArray),restcodes))
            MoveBarcodeCountLbl.config(text = countlbl)

            MoveRunDivisionBtn.configure(state = NORMAL)
            
            
            
    else:
        MoveRunDivisionBtn.configure(state = DISABLED)
        
        
        










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


if __name__ == '__main__':
    main()
