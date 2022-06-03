from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

from PyPDF2 import PdfFileReader
from pathlib import Path

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

global InputFile

def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('240x130+{}+{}'.format(scrnw, scrnh))
        
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
        global SLbl
        Lbl = Label(text="Папка проекта:", background="white")
        Lbl.place(x=16, y=10)
        
        global InputFileEntry
        InputFileEntry = Entry(fg="black", bg="white", width=20)
        InputFileEntry.place(x=20, y=32)
        
        global InputFileBtn
        InputFileBtn = Button(text='Выбор', command=SelectFile)
        #InputDirBtn.place(x=20, y=54, height=20)
        InputFileBtn.place(x=150, y=31, height=20)
        
        global RunBtn
        RunBtn = Button(text='Run', command=DPI)
        #InputDirBtn.place(x=20, y=54, height=20)
        RunBtn.place(x=20, y=55, height=20)

def BtnCmd():
    print('btn pressed')


def SelectFile():
    global InputFile

    InputFile = ""
    InputFile = filedialog.askopenfilename(title='Выберите файл на обработку', filetypes=(('PDF document', 'pdf'),))
    if InputFile:
        InputFileEntry.configure(state = NORMAL)
        InputFileEntry.delete(0,END)
        InputFileEntry.insert(0,str(InputFile))
        InputFileEntry.configure(state = DISABLED)
        print('IDC: InputFile : {0}'.format(InputFile))
    else:
        print('IDC: InputFile not selected')

def DPI():
    with open(InputFile, 'rb') as file:
        inp_pdf = PdfFileReader(file)
        media_box = inp_pdf.getPage(0).mediaBox

        min_pt = media_box.lowerLeft
        max_pt = media_box.upperRight
        pdf_width = max_pt[0]-min_pt[0]
        pdf_height = max_pt[1]-min_pt[1]

        print('DPI: height-{0}, width-{1}'.format(pdf_height, pdf_width))

if __name__ == '__main__':
    main()
