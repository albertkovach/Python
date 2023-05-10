from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

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
        
        global InputDirEntry
        InputDirEntry = Entry(fg="black", bg="white", width=20)
        InputDirEntry.place(x=20, y=32)
        
        global InputDirBtn
        InputDirBtn = Button(text='Выбор', command=SelectDir)
        #InputDirBtn.place(x=20, y=54, height=20)
        InputDirBtn.place(x=150, y=31, height=20)

def BtnCmd():
    print('btn pressed')


def SelectDir():
    global InputFile
    global IsInputSel

    InputFile = filedialog.askopenfilename(title='Выберите файл', filetypes=(('====', 'xml'),))
    if InputFile:
        InputDirEntry.configure(state = NORMAL)
        InputDirEntry.delete(0,END)
        InputDirEntry.insert(0,str(DivideInputFile))
        InputDirEntry.configure(state = DISABLED)
        print('IFC: InputFile :', DivideInputFile)
        
        IsInputSel = True
    else:
        print('IFC: InputFile not selected')


if __name__ == '__main__':
    main()
