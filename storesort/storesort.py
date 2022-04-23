from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

import os, sys
from pathlib import Path

global filename

from pdfminer.layout import LAParams, LTTextBox, LTText, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator


def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('300x130+{}+{}'.format(scrnw, scrnh))
        
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
        global InputDirLbl
        InputDirLbl = Label(text="Выберите папку с файлами:", background="white")
        InputDirLbl.place(x=16, y=10)
        
        global InputDirEntry
        InputDirEntry = Entry(fg="black", bg="white", width=20)
        InputDirEntry.place(x=20, y=32, width=180)
        InputDirEntry.configure(state = DISABLED)
        
        global InputDirBtn
        InputDirBtn = Button(text='Выбор', command=InputFileChoose)
        InputDirBtn.place(x=210, y=30, height=20)
        
        global InputFilesLbl
        InputFilesLbl = Label(text="", background="white")
        InputFilesLbl.place(x=16, y=50)
        
        
        global ExecuteBtn
        ExecuteBtn = Button(text='Запуск', command=Searcher)
        ExecuteBtn.place(x=20, y=80, width=60, height=20)





def InputFileChoose():
    global inputdir
    global FilesArray

    FilesArray = []
    inputdir = filedialog.askdirectory(title="Выбрать папку")
    
    if inputdir:
        InputDirEntry.configure(state = NORMAL)
        InputDirEntry.delete(0,END)
        InputDirEntry.insert(0,str(Path(inputdir).name))
        InputDirEntry.configure(state = DISABLED)
        
        print('inputdir:', inputdir)
        
        FilesArray.clear()
        for file in os.listdir(inputdir):
            if file.endswith(".pdf"):
                FilesArray.append(os.path.join(inputdir, file))
        ValidFilesCount = len(FilesArray)
        InputFilesLbl.config(text = 'Количество файлов PDF: ' + str(ValidFilesCount))
        print('number of valid files -', str(ValidFilesCount))
    else:
        print('inputdir is NOT defined')


def Searcher():
    global inputdir
    global FilesArray
    
    for f in range(len(FilesArray)):
        invkpp = FileTextSearch(FilesArray[f])
        print('Документ №{0}: {1}'.format(f+1, Path(FilesArray[f]).name))
        print('КПП документа: {0}'.format(invkpp))

    #isinputdir = os.path.isdir(inputdir)
    
def FolderWork(folder):
    print('inputdir is NOT defined')


def FileTextSearch(file):

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
                        text = textbox.get_text().replace('\n', '')
                        if len(text) == 22:
                            invoicekpp = text.partition("/")[2]
                            #print("_________________________")
                            #print("coord: {0}, {1}".format(textbox.bbox[0], textbox.bbox[1]))
                            #print("text: {0}".format(text))
                            #print("length: {0}".format(len(text)))
                            #print('КПП документа: {0}'.format(invoicekpp))
                            return invoicekpp
                            break
        return "NONE"




if __name__ == '__main__':
    main()
