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
    root.geometry('170x130+{}+{}'.format(scrnw, scrnh))
        
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
        Lbl = Label(text="Label", background="white")
        Lbl.place(x=16, y=10)
        
        global SEntry
        SEntry = Entry(fg="black", bg="white", width=20)
        SEntry.place(x=20, y=32)
        
        global SBtn
        SBtn = Button(text='Button', command=BtnCmd)
        SBtn.place(x=20, y=54, height=20)

def BtnCmd():
    print('btn pressed')

if __name__ == '__main__':
    main()
