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
        
        global InputDirBtn
        InputDirBtn = Button(text='Выбор', command=InputFileChoose)
        InputDirBtn.place(x=210, y=30, height=20)

def InputFileChoose():
    global filename
    filename = filedialog.askopenfilename(filetypes=(('', 'pdf'),))
    TextLbl.config(text = str(filename))
    print('Selected:', filename)

if __name__ == '__main__':
    main()
