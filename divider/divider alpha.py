from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

from threading import Thread
from multiprocessing import Process, Queue, current_process

import time

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

global MaxProcessCount
MaxProcessCount = 4

global DivideIsRunning
global DivideStartedTime

def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('400x170+{}+{}'.format(scrnw, scrnh))
        
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
        SBtn2.place(x=20, y=130, width=80, height=20)
        
        global thrlbl
        thrlbl = Label(text="Thread", background="white")
        thrlbl.place(x=110, y=20)
        
        global timelbl
        timelbl = Label(text="time", background="white")
        timelbl.place(x=230, y=20)
        
        
        global process1lbl
        process1lbl = Label(text="Process 1", background="white")
        process1lbl.place(x=20, y=45)
        
        global process2lbl
        process2lbl = Label(text="Process 2", background="white")
        process2lbl.place(x=20, y=65)
        
        global process3lbl
        process3lbl = Label(text="Process 3", background="white")
        process3lbl.place(x=20, y=85)
        
        global process4lbl
        process4lbl = Label(text="Process 3", background="white")
        process4lbl.place(x=20, y=105)



def BtnCmd():
    print('btn pressed')
    sthread = Thread(target=ProcessManager)
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
    
    

def ProcessManager():
    global DivideInputFilesArray
    global DivideOutputDir
    
    global MaxProcessCount
    
    global DivideIsRunning
    global DivideStartedTime
    
    DivideIsRunning = True
    DivideStartedTime = time.time()


    # Emulator vars init

    
    DivideOutputDir = "output"
    filearrayleng = 8
    DivideInputFilesArray = []
    MaxProcessCount = 3
    
### process just started var

    # ProcessManager:
    runningprocessescount = 0
    fileresultarray = []
    dataq = Queue()


    for i in range (filearrayleng):
        DivideInputFilesArray.append(i) ## DELETE THIS
        fileresultarray.append(2) # 2 -> awaiting, 1 -> in work, 0 -> ready
    print('ProcessManager: fileresultarray: {0}'.format(fileresultarray))


    processguinumarray = []
    #for i in range (MaxProcessCount):
    #    processguinumarray.append(0)
        
    firstloop = True
    while True:
        #print('ProcessManager: Loop started')
        
        # Checking if data received
        if not dataq.empty():
            # dataq - processnum, status, message
            processresult = dataq.get()
            print('ProcessManager: data received: {0}'.format(processresult))

            # Refreshing GUI with received data
            for g in range (len(processguinumarray)):
                if processguinumarray[g] == processresult[0]:
                    processguinum = g
                    #print('ProcessManager: processguinum: {0}'.format(processguinum))
                    print('ProcessManager: processguinumarray: {0}'.format(processguinumarray))
                    break

            if processguinum == 0:
                #process1lbl.config(text = str(processresult[2]))
                process1lbl.config(text = str(processresult))
            elif processguinum == 1:
                #process2lbl.config(text = str(processresult[2]))
                process2lbl.config(text = str(processresult))
            elif processguinum == 2:
                #process3lbl.config(text = str(processresult[2]))
                process3lbl.config(text = str(processresult))
            elif processguinum == 3:
                #process4lbl.config(text = str(processresult[2]))
                process4lbl.config(text = str(processresult))
                
            if processresult[1] == 0: # process terminated
                runningprocessescount = runningprocessescount - 1
                print('ProcessManager: process {0} terminated'.format(processresult[0]))
                processjustterminated = True
                
                # Resort guinumarray, because of process end
                filesawaiting = 0
                for u in range (len(fileresultarray)): # Scanning for awaiting files
                    if fileresultarray[u] == 2:
                        filesawaiting = filesawaiting + 1
                if not filesawaiting <= MaxProcessCount: # Checking if there are last 4 processes
                    processguinumarray[processguinum] = 0
                    processguinumtemparray = []
                    for t in range (len(processguinumarray)):
                        if processguinumarray[t] != 0:
                            processguinumtemparray.append(processguinumarray[t])
                    processguinumarray.clear()
                    for h in range (len(processguinumtemparray)):
                        processguinumarray.append(processguinumtemparray[h])
                    
                # Writing status of file in resultarray
                fileresultarray[processresult[0]] = 0
        
        # Checking amount of files that remains for exec
        filesawaiting = 0
        for r in range (len(fileresultarray)): # Scanning for awaiting files
            if fileresultarray[r] == 2:
                filesawaiting = filesawaiting + 1
        
        # Checking space for new process start
        if runningprocessescount < MaxProcessCount or firstloop == True:
            print('ProcessManager: runningprocessescount: {0}, MaxProcessCount: {1}'.format(runningprocessescount, MaxProcessCount))
            if firstloop:
                firstloop = False
                processguinumarray.clear()
                processestostart = MaxProcessCount
                print('ProcessManager: first loop !')
                print('ProcessManager: processestostart: {0}'.format(processestostart))
            elif runningprocessescount < filesawaiting and filesawaiting != 0:
                print('ProcessManager: more files awaiting than MaxProcessCount')
                print('ProcessManager: processestostart: {0}'.format(processestostart))
                processestostart = runningprocessescount
            elif runningprocessescount > filesawaiting and filesawaiting != 0:
                print('ProcessManager: runningprocessescount > filesawaiting')
                print('ProcessManager: processestostart: {0}'.format(processestostart))
                processestostart = filesawaiting
            
            if processestostart != 0:
                # Iterations for process start
                for s in range (processestostart+1):
                    for j in range (len(fileresultarray)):
                        if fileresultarray[j] == 2: # Scanning for awaiting files
                            fileresultarray[j] = 1 # Mark this file to "1 -> in work"
                            runningprocessescount = runningprocessescount + 1
                            print('ProcessManager: runningprocessescount: {0}'.format(runningprocessescount))
                            processguinumarray.append(j)
                            creatingprocessfilenum = j
                            print('ProcessManager: processguinumarray: {0}'.format(processguinumarray))
                            print('ProcessManager: creatingprocessfilenum: {0}'.format(creatingprocessfilenum))
                            break
                    # DividerEm(processnum, dataq, inputfile, outputdir)
                    subprocess = Process(target=DividerEm, args=(creatingprocessfilenum, dataq, DivideInputFilesArray[creatingprocessfilenum], DivideOutputDir))
                    subprocess.start()
        
        # Checking if all is done
        #print('ProcessManager: end fileresultarray: {0}'.format(fileresultarray))
        filesready = 0
        for e in range (len(fileresultarray)): # Scanning for ready files
            if fileresultarray[e] == 0:
                filesready = filesready + 1
        if filesready == len(fileresultarray):
            print('ProcessManager: All done !!!')
            break
                

        # чтение dataq если не пустой
            # запись в gui
            # if status == 0:
                # по procnum записать состояние "0 -> ready" в fileresultarray
                # уменьшить runningprocessescount
                
        # просканить fileresultarray на "2 -> awaiting" в filesawaiting
        
        # if runningprocessescount < MaxProcessCount
            # if runningprocessescount < filesawaiting ->> iterator = runningprocessescount
            #   else iterator = filesawaiting
                # for iterator на запуск процессов
                    # for, пробежать fileresultarray, найти первый с "2 -> awaiting"
                    # запуск процесса
                    # взвести его в "1 -> in work"
                    # break

        # просканить fileresultarray на "0 -> ready" в filesready
        # if filesready = len(fileresultarray)
            # break
            

    DivideIsRunning = False




def DividerEm(processnum, dataq, inputfile, outputdir):
    processname = current_process().name
    print("==== Divider №{0}--{1} STARTED: {2}, {3}".format(processnum, processname, inputfile, outputdir))
    
    length = 5
    for i in range(length):
        #print("== {0} exec data: {1}".format(processnum, i))
        message = "Поиск первых страниц... Чтение {0} из {1}, найдено: {2}".format(i, length, i)
        dataq.put([processnum, '1', message])
        time.sleep(0.2)
        
    dataq.put([processnum, 0, "Обработка завершена !"])
    
    print("==== Divider №{0}--{1} ENDED".format(processnum, processname))

        





def DivideTimeUpdater():
    global DivideIsRunning
    global DivideStartedTime

    while DivideIsRunning:
        result = time.time() - DivideStartedTime
        result = datetime2.timedelta(seconds=round(result))
        timelbl.config(text = str(result))
        time.sleep(0.01)




if __name__ == '__main__':
    main()
