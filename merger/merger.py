from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from PyPDF2 import PdfFileMerger
from threading import Thread
import os

 
class Example(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.initUI()
    
    def initUI(self):
        self.parent.title("Merge PDF")
        self.pack(fill=BOTH, expand=1)

        fx = 20
        fy = 40
        
        global DirPathEntry
        DirPathEntry = Entry(fg="black", bg="white", width=40)
        DirPathEntry.place(x=fx, y=fy)
        
        global DirChooseBtn
        DirChooseBtn = Button(text='Выбор', command=DirChoose)
        DirChooseBtn.place(x=fx+250, y=fy-1, height=20)

        global DirCalcBtn
        DirCalcBtn = Button(text='Посчитать', command=CountFiles)
        DirCalcBtn.place(x=fx, y=fy+22, height=20)

        global FilesCountLbl
        FilesCountLbl = Label(text="", background="white")
        FilesCountLbl.place(x=fx+75, y=fy+22)

        global MergeBtn
        MergeBtn = Button(text="Выполнить слияние", command=StartMergingThread)
        MergeBtn.place(x=fx, y=fy+52, height=20)

        global directory
        global FilesArray
        FilesArray = []


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
    
    app = Example(root)
    root.mainloop()


def DirChoose(): 
    directory = filedialog.askdirectory(title="Выбрать папку") #initialdir=""
    if directory:
        DirPathEntry.delete(0,END)
        DirPathEntry.insert(0,directory)
        print(directory)
        CountFiles()


def CountFiles():
    directory = DirPathEntry.get()
    isDirectory = os.path.isdir(directory)
    if isDirectory:
        FilesArray.clear()
        print('Path is valid: ', isDirectory)
        for file in os.listdir(directory):
            if file.endswith(".pdf"):
                FilesArray.append(os.path.join(directory, file))
        FilesCountLbl.config(text = 'Количество файлов PDF: ' + str(len(FilesArray)))
        print('Number of valid files: ', len(FilesArray))
    else:
        FilesCountLbl.config(text = 'Неправильный путь')

def StartMergingThread():
    directory = DirPathEntry.get()
    isDirectory = os.path.isdir(directory)
    if isDirectory:
        CountFiles()
        mergethread = Thread(target=PDFmerge)
        mergethread.start()
        mergethread.join()

def PDFmerge():
    pdfmerger = PdfFileMerger()
    for i in range(0, len(FilesArray)):
        pdfmerger.append(FilesArray[i])
        print(i)
        print(FilesArray[i])
    pdfmerger.write("C:\\Users\\ORB User\\Desktop\\pdf\\result.pdf")

if __name__ == '__main__':
    main()

