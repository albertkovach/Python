from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

import PyPDF2
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter
from PyPDF2 import PdfFileMerger

import textract, re


global InputFile
global WordsCount


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
    root.geometry('280x130+{}+{}'.format(scrnw, scrnh))
        
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
        global MainLbl
        MainLbl = Label(text="Выберите файл:", background="white")
        MainLbl.place(x=16, y=10)
        
        global InputFileEntry
        InputFileEntry = Entry(fg="black", bg="white", width=30)
        InputFileEntry.place(x=20, y=32)
        
        global SelectFileBtn
        SelectFileBtn = Button(text='Выбор', command=SelectFile)
        SelectFileBtn.place(x=210, y=30, height=20)
        
        global CountPagesBtn
        CountPagesBtn = Button(text='Посчитать', command=CountWords)
        CountPagesBtn.place(x=190, y=55, height=20)
        
        global PagesCountLbl
        PagesCountLbl = Label(text="", background="white")
        PagesCountLbl.place(x=16, y=52)
        
        

def SelectFile():
    print('SelectFile pressed')
    
    global InputFile

    InputFile = filedialog.askopenfilename(title='Выберите файл на обработку', filetypes=(('PDF document', 'pdf'),))
    if InputFile:
        InputFileEntry.configure(state = NORMAL)
        InputFileEntry.delete(0,END)
        InputFileEntry.insert(0,str(InputFile))
        InputFileEntry.configure(state = DISABLED)
        print('IFC: InputFile :', InputFile)
        PagesCount = PDFCountPages(InputFile)
        PagesCountLbl.config(text = 'Страниц в файле: ' + str(PagesCount))

    else:
        print('IFC: InputFile not selected')
        
        
def CountWords(): 
    print('CountWords pressed')
    

    text = textract.process(InputFile)
    words = re.findall(r"[^\W_]+", text, re.MULTILINE)
    print(len(words))
    print(words)
    
    
    
        
        
def PDFCountPages(inputfile):
    pdf = PdfFileReader(inputfile)
    pagecount = pdf.getNumPages()
    pdf = ""
    return pagecount
    
    

if __name__ == '__main__':
    main()
