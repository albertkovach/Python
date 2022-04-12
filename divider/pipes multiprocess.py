from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

from threading import Thread
from multiprocessing import Process, Pipe, current_process

import time

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

global toDivProcess1Pipe, DivProcess1Pipe
global toDivProcess2Pipe, DivProcess2Pipe
global toDivProcess3Pipe, DivProcess3Pipe
global spprocess1

def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('400x150+{}+{}'.format(scrnw, scrnh))
        
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
        
        global SBtn
        SBtn = Button(text='Button', command=BtnCmd)
        SBtn.place(x=20, y=20, width=80, height=20)
        
        global SBtn2
        SBtn2 = Button(text='Button2', command=Btn2Cmd)
        SBtn2.place(x=20, y=110, width=80, height=20)
        
        global thrlbl
        thrlbl = Label(text="Thread", background="white")
        thrlbl.place(x=110, y=20)
        
        global thrstatlbl
        thrstatlbl = Label(text="time", background="white")
        thrstatlbl.place(x=230, y=20)
        
        
        global plbl1
        plbl1 = Label(text="Process 1", background="white")
        plbl1.place(x=20, y=45)
        
        global plbl2
        plbl2 = Label(text="Process 2", background="white")
        plbl2.place(x=20, y=65)
        
        global plbl3
        plbl3 = Label(text="Process 3", background="white")
        plbl3.place(x=20, y=85)



def BtnCmd():
    print('btn pressed')
    sthread = Thread(target=ThreadPM)
    sthread.start()



def Btn2Cmd():
    global spprocess1
    
    currtime = time.time()
    currtimelbl = "{0}, {1}".format(bool(spprocess1.is_alive), currtime)
    thrstatlbl.config(text = currtimelbl)
    
    try:
        result1 = toDivProcess1Pipe.recv()
        print(">>>>> btn2 pressed: result1[1]: {0}, prcs: {1}, alive: {2}".format(result1[1], spprocess1, bool(spprocess1.is_alive)))
        print(">>>>> btn2 pressed: full alive status: {0}".format(spprocess1.is_alive))
        
    except:
        print(">>>>> btn2 pressed: cant read pipe")
    
    

def ThreadPM():
    print('**** ThreadPM Started ****')
    thrlbl.config(text = "ThreadPM Started")
    
    global toDivProcess1Pipe
    global DivProcess1Pipe
    global spprocess1
    
    toDivProcess1Pipe, DivProcess1Pipe= Pipe()
    spprocess1 = Process(target=ParallelProcess, args=(DivProcess1Pipe,))
    spprocess1.start()
    
    



    toDivProcess1Pipe.send([0])
    # Zero-iteration and define
    result1 = toDivProcess1Pipe.recv()
    print("**** ThreadPM: result1[0]: {0}, result1[1]: {1}".format(result1[0], result1[1]))

    # Until process end message
    while int(result1[0]) != 0:
        result1 = toDivProcess1Pipe.recv()
        plbl1.config(text = result1)
        print("**** ThreadPM: result1[0]: {0}, result1[1]: {1}".format(result1[0], result1[1]))
        
        currtime = time.time()
        currtimelbl = (currtime)
        thrstatlbl.config(text = currtimelbl)
    
    # Unusual check
    if int(result1[0]) == 0:
        print('**** ThreadPM: Process ended computations')
        toDivProcess1Pipe.send([30])


    print('**** ThreadPM: result1[0]: {0}, result1[1]: {1}, prcs {2}'.format(result1[1], result1[1], spprocess1))
    
    thrlbl.config(text = "ThreadPM Ended")
    print('**** ThreadPM Ended ****')



def ParallelProcess(pipe):
    proc_name = current_process().name
    print("========== ParallelProcess {0} started".format(proc_name))
    
    data = pipe.recv()
    print("{0} Recieved data: {1}".format(proc_name, data))
    
    num = data[0]
    length = 50
    
    for i in range(length):
        num = num + 1
        #print("{0} exec data: {1}".format(proc_name, num))
        message = "Поиск первых страниц... Чтение {0} из {1}, найдено: {2}".format(i, length, i)
        pipe.send(['1', message])
        time.sleep(0.2)
        
    pipe.send([0, "Обработка завершена !"])
    #pipe.close()
    
    print("==========  ParallelProcess {0} ended".format(proc_name))

        










if __name__ == '__main__':
    main()
