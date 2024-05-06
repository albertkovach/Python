from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from PIL import Image, ImageTk

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
    root.geometry('800x800+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()

class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)   
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
        
        
        imagefile = Image.open("C:\\Users\\user\\Documents\\GitHub\\Python\\rasremote\\castle.png")
        pilimage = ImageTk.PhotoImage(imagefile)

        # ttk.Label(self, image=self.python_image).pack()
        # image = PhotoImage(file="C:\\Users\\user\\Documents\\GitHub\\Python\\rasremote\\castle.png")
        # img = Label(root, image=pilimage)
        # img.place(x=1, y=1, width=200, height=200)
        
        
        global SImage
        SImage = Label(text='Button', image=pilimage)
        SImage.image = pilimage
        SImage.place(x=80, y=80)
        
        
        
        
        
        

def BtnCmd():
    print('btn pressed')

if __name__ == '__main__':
    main()
