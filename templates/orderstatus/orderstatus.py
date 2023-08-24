from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
import tkinter.font as TkFont
import time

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150



def main():
    global root

    root = Tk()
    root.resizable(False, False)
    
    root.bind('<KeyRelease>', lambda x: SearchDef()) #Return
   
    style=ttk.Style()
    style.theme_use('classic')
    style.configure("Vertical.TScrollbar", background="green", bordercolor="red", arrowcolor="white")
   
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('345x480+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()



class GUI(Frame):
    def __init__(self, parent):  
        Frame.__init__(self, parent)   # background='#f5f5f5'
        self.parent = parent
        self.parent.title("")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        listfont= TkFont.Font(family='Helvetica', size=16, weight='bold')
        btnfontsml= TkFont.Font(family='Helvetica', size=14, weight='bold')
        btnfontbig= TkFont.Font(family='Helvetica', size=20, weight='bold')
        
        global activesearch
        activesearch = 0
        
        global activefield
        activefield = 0
        
        global InWorkLbl
        global InWorkArray
        global InWorkList
        global InWorkListbox
        global InWorkSearch
        global InWorkSearchEntry 
        
        global ReadyLbl
        global ReadyArray
        global ReadyList
        global ReadyListbox
        global ReadySearchEntry
        
        ofswx = 15
        ofswy = 10
        
        InWorkArray = []
        startx=5000
        for x in range(5):
            InWorkArray.append(startx+x)
        
        InWorkLbl = Label(text="В работе:", font=listfont)
        InWorkLbl.place(x=ofswx, y=ofswy)
        
        InWorkListbox = Listbox(width=8, height=15, font=listfont)
        for x in InWorkArray:
            InWorkListbox.insert(END, x)
        InWorkListbox.place(x=ofswx+6, y=ofswy+35)
        InWorkListbox.bind("<FocusIn>", InWorkListboxFocIn)
        InWorkListbox.bind("<FocusOut>", InWorkListboxFocOut)
        scroll = Scrollbar(command=InWorkListbox.yview)
        scroll.place(x=ofswx+105, y=ofswy+35, width=20, height=379)
        InWorkListbox.config(yscrollcommand=scroll.set)
        InWorkSearchEntry = Entry(validate="focus", validatecommand=InWorkSearchDef, font=listfont)
        InWorkSearchEntry.place(x=ofswx+6, y=ofswy+420, width=100)
        InWorkListboxColorize()

        ofsrx = 200
        ofsry = 10
        
        ReadyArray = []
        startx=1000
        for x in range(3):
            ReadyArray.append(startx+x)
        
        ReadyLbl = Label(text="Готово:", font=listfont)
        ReadyLbl.place(x=ofsrx, y=ofsry)
        
        ReadyListbox = Listbox(width=8, height=15, font=listfont, fg='#757575')
        for x in ReadyArray:
            ReadyListbox.insert(END, x)
        ReadyListbox.place(x=ofsrx+6, y=ofsry+35)
        scroll = Scrollbar(command=ReadyListbox.yview)
        scroll.place(x=ofsrx+105, y=ofsry+35, width=20, height=379)
        ReadyListbox.config(yscrollcommand=scroll.set)
        ReadySearchEntry = Entry(validate="focus", validatecommand=ReadySearchDef, font=listfont)
        ReadySearchEntry.place(x=ofsrx+6, y=ofsry+420, width=100)
        ReadyListboxColorize()
        
        
        global AddNewBtn
        AddNewBtn = Button(text='+', command=AddNew, font=btnfontbig, state = DISABLED, bg='#b5b5b5')
        AddNewBtn.place(x=130, y=429, width=30, height=29)
        
        global DeleteBtn
        DeleteBtn = Button(text='×', command=Delete, font=btnfontbig, state = DISABLED, bg='#b5b5b5')
        DeleteBtn.place(x=145, y=376, width=30, height=30)

        
        global MoveToReadyBtn
        MoveToReadyBtn = Button(text='>>', command=MoveToReady, font=btnfontsml, fg='#000000', bg='#f8fcbb') #
        MoveToReadyBtn.place(x=165, y=170, width=30, height=30)
        
        global MoveToInWorkBtn
        MoveToInWorkBtn = Button(text='<<', command=MoveToInWork, font=btnfontsml, fg='#000000', bg='#dfeef5') # d9d9d9
        MoveToInWorkBtn.place(x=145, y=260, width=30, height=30)





    ###### Обработчики кнопок

def MoveToReady():
    global InWorkArray
    global InWorkListbox
    global InWorkSearchEntry

    global ReadyArray
    global ReadyListbox
    global ReadySearchEntry

    for i in InWorkListbox.curselection():
        ReadyArray.append(InWorkListbox.get(i))
        InWorkArray.remove(InWorkListbox.get(i))
        
    if InWorkSearchEntry.get() == '':
        InWorkListbox.delete(0, END)
        for x in InWorkArray:
            InWorkListbox.insert(END, x)
    else:
        InWorkListbox.delete(0, END)
        texttosearch = InWorkSearchEntry.get()
        for item in InWorkArray:
            if str(texttosearch) in str(item):
                if str(item).index(texttosearch) == 0:
                    InWorkListbox.insert(END, item)
    
    InWorkListboxColorize()
    
    ReadySearchEntry.delete(0, END)
    ReadySearchEntry.insert(0, '')
    
    ReadyListbox.delete(0, END)
    for x in ReadyArray:
        ReadyListbox.insert(END, x)
    
    ReadyListboxColorize()


def MoveToInWork():
    global InWorkArray
    global InWorkListbox
    global InWorkSearchEntry

    global ReadyArray
    global ReadyListbox
    global ReadySearchEntry
    
    answer = messagebox.askyesno(title='Подтверждение', message='Уверены, что хотите вернуть из готовых ?')
    if answer:

        for i in ReadyListbox.curselection():
            InWorkArray.append(ReadyListbox.get(i))
            ReadyArray.remove(ReadyListbox.get(i))
            
        if ReadySearchEntry.get() == '':
            ReadyListbox.delete(0, END)
            for x in ReadyArray:
                ReadyListbox.insert(END, x)
        else:
            ReadyListbox.delete(0, END)
            texttosearch = ReadySearchEntry.get()
            for item in ReadyArray:
                if str(texttosearch) in str(item):
                    if str(item).index(texttosearch) == 0:
                        ReadyListbox.insert(END, item)
        
        ReadyListboxColorize()
        
        InWorkSearchEntry.delete(0, END)
        InWorkSearchEntry.insert(0, '')
        
        InWorkListbox.delete(0, END)
        for x in InWorkArray:
            InWorkListbox.insert(END, x)
        
        InWorkListboxColorize()
    

def AddNew():
    global InWorkArray
    global InWorkListbox
    global InWorkSearchEntry
    global ReadyArray
    global activesearch
    
    if activesearch == 1:
        texttoadd = InWorkSearchEntry.get()

        for item in InWorkArray:
            if str(texttoadd) == str(item):
                messagebox.showerror("", "Такое уже есть в работе !")
                return
        
        for item in ReadyArray:
            if str(texttoadd) == str(item):
                messagebox.showerror("", "Такое уже есть в готовых !")
                return

        InWorkArray.append(texttoadd)
        InWorkListbox.delete(0, END)
        SearchDef()


def Delete():
    global InWorkArray
    global InWorkListbox
    global InWorkSearchEntry
    global activesearch
    
    answer = messagebox.askyesno(title='Подтверждение', message='Уверены, что хотите удалить ?')
    if answer:

        for i in InWorkListbox.curselection():
            InWorkArray.remove(InWorkListbox.get(i))
        
        if InWorkSearchEntry.get() == '':
            InWorkListbox.delete(0, END)
            for x in InWorkArray:
                InWorkListbox.insert(END, x)
        else:
            InWorkListbox.delete(0, END)
            texttosearch = InWorkSearchEntry.get()
            for item in InWorkArray:
                if str(texttosearch) in str(item):
                    if str(item).index(texttosearch) == 0:
                        InWorkListbox.insert(END, item)
        
        InWorkListboxColorize()


def SearchDef():
    global InWorkArray
    global InWorkListbox
    global InWorkSearchEntry

    global ReadyArray
    global ReadyListbox
    global ReadySearchEntry

    global activesearch
    
    if activesearch == 1:
        texttosearch = InWorkSearchEntry.get()
        
        if InWorkSearchEntry.get() == '':
            InWorkListbox.delete(0, END)
            for x in InWorkArray:
                InWorkListbox.insert(END, x)
        else:
            InWorkListbox.delete(0, END)
            for item in InWorkArray:
                if str(texttosearch) in str(item):
                    if str(item).index(texttosearch) == 0:
                        InWorkListbox.insert(END, item)
        InWorkListboxColorize()
    
    elif activesearch == 2:
        texttosearch = ReadySearchEntry.get()
        
        if ReadySearchEntry.get() == '':
            ReadyListbox.delete(0, END)
            for x in ReadyArray:
                ReadyListbox.insert(END, x)
        else:
            ReadyListbox.delete(0, END)
            for item in ReadyArray:
                if str(texttosearch) in str(item):
                    if str(item).index(texttosearch) == 0:
                        ReadyListbox.insert(END, item)
        ReadyListboxColorize()
    return True




    ###### Раскраска   
    
def InWorkListboxColorize():
    #for i in range(0,InWorkListbox.size(),2):
    #    InWorkListbox.itemconfigure(i, background='#e6e6e6')
    for i in range(InWorkListbox.size()):
        InWorkListbox.itemconfigure(i, background='#c1f7c5')   
  
def ReadyListboxColorize():
    #for i in range(0,ReadyListbox.size(),2):
    for i in range(ReadyListbox.size()):
        ReadyListbox.itemconfigure(i, background='#e6e6e6')





    ###### Обработчики фокуса

def InWorkListboxFocIn(event):
    global DeleteBtn
    DeleteBtn.configure(state = NORMAL, bg='#ff4278')
    return True
    
def InWorkListboxFocOut(event):
    global DeleteBtn
    DeleteBtn.configure(state = DISABLED, bg='#b5b5b5')

def InWorkSearchDef():
    global activesearch
    global AddNewBtn
    if activesearch == 1:
        activesearch = 0
        AddNewBtn.configure(state = DISABLED, bg='#b5b5b5')
    else:
        activesearch = 1
        AddNewBtn.configure(state = NORMAL, bg='#42ff91')
    return True
    
def ReadySearchDef():
    global activesearch
    if activesearch == 2:
        activesearch = 0
    else:
        activesearch = 2
    return True  

    
    







if __name__ == '__main__':
    main()
