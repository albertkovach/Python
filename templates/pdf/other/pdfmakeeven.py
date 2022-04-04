from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

global filename
import PyPDF2
from pathlib import Path

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
    root.geometry('250x90+{}+{}'.format(scrnw, scrnh))
        
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
        global TextLbl
        TextLbl = Label(text="", background="white")
        TextLbl.place(x=20, y=45)
        
        global ChooseFileBtn
        ChooseFileBtn = Button(text='Выбрать', command=BtnCmd1)
        ChooseFileBtn.place(x=20, y=20, height=20)

        global RunBtn
        RunBtn = Button(text='Запуск', command=BtnCmd2)
        RunBtn.place(x=90, y=20, height=20)

def BtnCmd1():
    global filename
    filename = filedialog.askopenfilename()
    TextLbl.config(text = str(filename))
    print('Selected:', filename)


def BtnCmd2():
    global filename
    outputdir = Path(filename).parent.absolute()
    outputfilename = 'EXEC_' + Path(filename).stem + '.pdf'
    outputfile = Path(outputdir, outputfilename)
    print('outputfile:', outputfile)
    make_even_page(filename, outputfile)




def make_even_page(in_fpath, out_fpath):
    reader = PyPDF2.PdfFileReader(in_fpath)
    writer = PyPDF2.PdfFileWriter()
    for i in range(reader.getNumPages()):
        writer.addPage(reader.getPage(i))
    if reader.getNumPages() % 2 == 1:
        _, _, w, h = reader.getPage(0)['/MediaBox']
        writer.addBlankPage(w, h)
    with open(out_fpath, 'wb') as fd:
        writer.write(fd)


if __name__ == '__main__':
    main()
