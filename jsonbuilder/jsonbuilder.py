from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

import os, shutil
from datetime import datetime
from pathlib import Path
from PyPDF2 import PdfFileReader



def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('435x240+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()



class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.parent.title("JSON Builder")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        global InputFile
        global InputFolder
        global InputFilesArray
        InputFilesArray = []
        global FolderMode
        FolderMode = False
        global NoError
        NoError = True
        global OrbitaPath
        OrbitaPath = "C:\Orbita v.2\Scan"
        
    
        global InputFileLbl
        InputFileLbl = Label(text="Документ PDF:", background="white", font=("TkDefaultFont", 10, "bold"))
        InputFileLbl.place(x=16, y=12+10)
        
        global InputFileEntry
        InputFileEntry = Entry(fg="black", bg="white", width=25)
        InputFileEntry.place(x=120, y=12+10)
        
        global InputFileBtn
        InputFileBtn = Button(text='Выбор', width=7, command=SelectFile)
        InputFileBtn.place(x=285, y=11+10, height=20)
        
        global InputFolderBtn
        InputFolderBtn = Button(text='Папка', width=7, command=SelectFolder)
        InputFolderBtn.place(x=355, y=11+10, height=20)
        
        global InputFolderRefrBtn
        InputFolderRefrBtn = Button(text='Обновить', width=7, command=RefreshFolder)
        InputFolderRefrBtn.place(x=355, y=11+10+25, height=20)
        
        
        global InputFilePCLbl
        InputFilePCLbl = Label(text="", background="white")
        InputFilePCLbl.place(x=75, y=35+10)
        
        
        
        global FolderNameLbl
        FolderNameLbl = Label(text="Фолдер:", background="white")
        FolderNameLbl.place(x=16, y=85)
        
        global FolderNameEntry
        FolderNameEntry = Entry(fg="black", bg="white", width=18)
        FolderNameEntry.place(x=75, y=85)
        
        global BoxNameLbl
        BoxNameLbl = Label(text="Короб:", background="white")
        BoxNameLbl.place(x=205, y=85)
        
        global BoxNameEntry
        BoxNameEntry = Entry(fg="black", bg="white", width=18)
        BoxNameEntry.place(x=260, y=85)
        
        
        
        global ProjectIDLbl
        ProjectIDLbl = Label(text="ID:", background="white")
        ProjectIDLbl.place(x=340, y=120)
        
        global ProjectIDEntry
        ProjectIDEntry = Entry(fg="black", bg="white", width=5)
        ProjectIDEntry.place(x=370, y=120)
        
        global ProjectNameLbl
        ProjectNameLbl = Label(text="Имя проекта:", background="white")
        ProjectNameLbl.place(x=25, y=120)
        
        global ProjectNameEntry
        ProjectNameEntry = Entry(fg="black", bg="white", width=30)
        ProjectNameEntry.place(x=130, y=120)
        
        
        
        global UserIDLbl
        UserIDLbl = Label(text="ID:", background="white")
        UserIDLbl.place(x=340, y=155)
        
        global UserIDEntry
        UserIDEntry = Entry(fg="black", bg="white", width=5)
        UserIDEntry.place(x=370, y=155)
        
        global UserNameLbl
        UserNameLbl = Label(text="Имя оператора:", background="white")
        UserNameLbl.place(x=25, y=155)
        
        global UserNameEntry
        UserNameEntry = Entry(fg="black", bg="white", width=30)
        UserNameEntry.place(x=130, y=155)
        
        
        
        global RunBtn
        RunBtn = Button(text='Создать JSON и переместить', command=Run)
        RunBtn.place(x=19, y=195, width=230, height=25)
        
        global OpenOutputBtn
        OpenOutputBtn = Button(text='Открыть папку оператора', command=OpenOutputFolder)
        OpenOutputBtn.place(x=260, y=195, width=155, height=25)
        
        DefaultVarInit()




def SelectFile():
    global InputFile  
    global FolderMode  

    InputFile = ""
    InputFile = filedialog.askopenfilename(title='Выберите файл PDF:', filetypes=(('PDF files', 'pdf'),))
    
    if InputFile:
        print('InputFile : {0}'.format(InputFile))
        
        InputFileFullName = Path(InputFile).name
        InputFileName = InputFileFullName[:len(InputFileFullName) - 4]
        
        InputFileEntry.configure(state = NORMAL)
        InputFileEntry.delete(0,END)
        InputFileEntry.insert(0,str(InputFileName))
        InputFileEntry.configure(state = DISABLED)
        
        InputFilePC = CountPages(InputFile)
        InputFilePCLbl.config(text = "Количество страниц: {0}".format(InputFilePC))
        
        RunBtn.configure(state = NORMAL)
        
    else:
        print('InputFile not selected')
        
    FolderMode = False



def SelectFolder():
    global InputFolder
    global InputFilesArray
    global FolderMode

    InputFolder = ""
    InputFolder = filedialog.askdirectory(title='Выберите папку на обработку')
    
    if InputFolder:
        print('InputFolder : {0}'.format(InputFolder))
        InputFileEntry.configure(state = NORMAL)
        InputFileEntry.delete(0,END)
        InputFileEntry.insert(0,str(Path(InputFolder).name))
        InputFileEntry.configure(state = DISABLED)
        
        RefreshFolder()
        print('** Array is packed!\n')
        
        if len(InputFilesArray) > 0:
            RunBtn.configure(state = NORMAL)
            InputFolderRefrBtn.configure(state = NORMAL)
                
    else:
        print('InputFolder not selected')
        
    FolderMode = True



def RefreshFolder():
    global InputFolder
    global InputFilesArray
    
    print('** Updating folder...')
    InputFilesArray.clear()
    
    AllPageCount = 0

    for file in os.listdir(InputFolder):
        if file.endswith(".pdf"):
            InputFilePath = Path(InputFolder, file).as_posix()
            InputFilesArray.append(InputFilePath)
            print('*** File : {0}'.format(InputFilePath))
            AllPageCount = AllPageCount + CountPages(InputFilePath)
            
    InputFilePCLbl.config(text = 'Количество файлов: {0}, страниц: {1}'.format(len(InputFilesArray), AllPageCount))



def Run():
    global InputFile
    global InputFolder
    global InputFilesArray
    
    global FolderMode
    global NoError
    
    if FolderMode:
        for doc in range(len(InputFilesArray)):
            CreateJSON(InputFilesArray[doc])
    else:
        CreateJSON(InputFile)
        
    if NoError:
        messagebox.showinfo("", "Задача выполнена !")
        print('ALL DONE !\n')
    else:
        messagebox.showerror("", "Ошибка при выполнении !")
        print('ALL DONE !\n')



def CreateJSON(InputFile):
    global OrbitaPath
    global NoError
    
    InputFilePC = CountPages(InputFile)
    
    InputFileFullName = Path(InputFile).name
    InputFileName = InputFileFullName[:len(InputFileFullName) - 4]
    
    FolderName = FolderNameEntry.get()
    BoxName = BoxNameEntry.get()
    ProjectID = ProjectIDEntry.get()
    ProjectName = ProjectNameEntry.get()
    UserID = UserIDEntry.get()
    UserName = UserNameEntry.get()
    
    JSONfileName = InputFileName + ".json"
    JSONfile = Path(Path(InputFile).parent, JSONfileName)


    try:
        print('Creating JSONfile : {0}'.format(JSONfile))
        
        json = open(JSONfile, 'w', newline='')
        json.writelines('{\n')
        json.writelines('  "docName": "' + InputFileName + '",' + '\n')
        json.writelines('  "project": {' + '\n')
        json.writelines('    "id": ' + ProjectID + ',\n')
        json.writelines('    "name": "' + ProjectName + '"\n')
        json.writelines('  },\n')
        json.writelines('  "scanUser": {\n')
        json.writelines('    "id": ' + UserID + ',\n')
        json.writelines('    "name": "' + UserName + '"\n')
        json.writelines('  },\n')
        json.writelines('  "boxName": "' + BoxName + '",\n')
        if str(FolderName) == "null":
            json.writelines('  "scanFolderName": ' + FolderName + ',\n')
        else:
            json.writelines('  "scanFolderName": "' + FolderName + '",\n')
        now = datetime.now()
        date_time = now.strftime("%Y-%m-%dT%H:%M:%S.6528267+03:00")
        json.writelines('  "scanDTime": "' + date_time + '",\n')
        json.writelines('  "scanByFeeder": true,\n')
        json.writelines('  "platePages": null,\n')
        json.writelines('  "pageCount": ' + str(InputFilePC) + ',\n')
        json.writelines('  "pages": [\n')
        json.close()
        
        for b in range(InputFilePC):
            SinglePageBlock(JSONfile, b+1, InputFilePC)
            
        json = open(JSONfile, 'a', newline='')
        json.writelines('  ],\n')
        json.writelines('  "scanSize": {\n')
        json.writelines('    "A4": ' + str(InputFilePC) + '\n')
        json.writelines('  },\n')
        json.writelines('  "itemAttributes": null\n')
        json.writelines('}')
        json.close()
    except:
        messagebox.showerror("", "Невозможно создать .json !")
        NoError = False



    try:
        completedfilesdir = Path(OrbitaPath, UserID)
        revcompletedfilesdir = completedfilesdir.as_posix()
        ismovedirexist = os.path.exists(revcompletedfilesdir)
        print('revcompletedfilesdir : {0}'.format(revcompletedfilesdir))
        if not ismovedirexist:
            os.makedirs(revcompletedfilesdir)
            print('created !')
    except:
        messagebox.showerror("", "Невозможно создать папку !")
        NoError = False



    try:
        print('moving started !')
        InputFileOutPath = Path(completedfilesdir, Path(InputFile).name)
        RevisedInputFileOutPath = InputFileOutPath.as_posix()
        shutil.copyfile(InputFile, InputFileOutPath)
        
        JSONfileOutPath = Path(completedfilesdir, Path(JSONfile).name)
        RevisedJSONfileOutPath = JSONfileOutPath.as_posix()
        shutil.copyfile(JSONfile, RevisedJSONfileOutPath)
    except:
        messagebox.showerror("", "Невозможно скопировать файлы !")
        NoError = False



def SinglePageBlock(JSONfile, pagenumber, pagecount):
    global NoError
    
    try:
        json = open(JSONfile, 'a', newline='')
        json.writelines('    {\n')
        json.writelines('      "PageNumber": ' + str(pagenumber) + ',\n')
        json.writelines('      "ScanType": 0,\n')
        json.writelines('      "ScanSize": 0,\n')
        json.writelines('      "ScannerName": "PaperStream IP fi-7260",\n')
        json.writelines('      "ScannerSN": "A3MAD01397",\n')
        json.writelines('      "ComputerName": "DC038-PC"\n')
        if pagenumber == pagecount:
            json.writelines('    }\n')
        else:
            json.writelines('    },\n')
        json.close()
        
    except:
        messagebox.showerror("", "Невозможно создать .json !")
        NoError = False




def OpenOutputFolder():
    global OrbitaPath
    
    UserID = UserIDEntry.get()
    userfolder = Path(OrbitaPath, UserID)
    revuserfolder = userfolder.as_posix()
    
    isrevuserfolderexist = os.path.exists(revuserfolder)

    if not isrevuserfolderexist:
        os.makedirs(revuserfolder)
    
    os.startfile(revuserfolder)


def DefaultVarInit():
    FolderNameEntry.insert(0,"null")
    BoxNameEntry.insert(0,"2600000000000")
    ProjectIDEntry.insert(0,"4921")
    ProjectNameEntry.insert(0,"НАТС ЦОД")
    UserIDEntry.insert(0,"2111")
    UserNameEntry.insert(0,"Зарецкая Людмила Сергеевна")
    
    RunBtn.configure(state = DISABLED)
    InputFolderRefrBtn.configure(state = DISABLED)


def CountPages(file):
    pdf = PdfFileReader(file)
    try:
        pagecount = pdf.getNumPages()
    except:
        msgbxlbl = ['Ошибка открытия файла !', str(file)]
        messagebox.showerror("", "\n".join(msgbxlbl))
        return 0
    pdf = ""
    return pagecount

if __name__ == '__main__':
    main()
