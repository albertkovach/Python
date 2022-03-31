from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

import win32api
import win32print
import PyPDF2
global root
import os



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
    win32api.ShellExecute(0, "print", filename, None,  ".",  0)
    print('Sent to print')

    

if __name__ == '__main__':
    main()
