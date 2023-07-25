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
        global Lbl
        Lbl = Label(text="Текст", background="white")
        Lbl.place(x=16, y=10)
        
        global TextEntry
        TextEntry = Entry(fg="black", bg="white", width=20)
        TextEntry.place(x=20, y=35)
        
        global Comm1Btn
        Comm1Btn = Button(text='Comm1', command=Comm1)
        Comm1Btn.place(x=20, y=60, height=20)








def Comm1():
    print('Comm1 !')



if __name__ == '__main__':
    main()
