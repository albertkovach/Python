from ctypes import windll, byref
from ctypes.wintypes import SMALL_RECT
import sys
import os

import time
import datetime as datetime2
from datetime import datetime

import pyfiglet

from tqdm import tqdm
from tqdm.notebook import tqdm_notebook

from clint.textui import colored, puts
import six

import colorama
from termcolor import colored
    


global counternum
global starttime



def main():
    global counternum
    counternum = 1
    
    colorama.init()

    #os.system('color 1f')
    os.system('mode con: cols=75 lines=20')
    logo = pyfiglet.figlet_format("P D F  M i n e r ", font = "slant"  )
    print(logo)


    print("Обработка файла 6 из 15: 383.pdf")

    Counter2()

def Counter2():
    global counternum
    global starttime
    
    starttime = time.time()
    
    print("")
    print("")
    
    while True:
        if counternum == 1:
            for i in range (3333):
                text = 'Сортировка файлов: {0} из 3333'.format(i)
                PrintToScreen(text)
            counternum = 2
        else:
            for i in range (2222):
                text = 'ПЕРЕМЕЩЕНИЕ файлов: {0} из 2222'.format(i)
                PrintToScreen(text)
            counternum == 1
    
    
    

def Counter():
#print('Сортировка файлов: {0} из 200000'.format(i),end='\r')
    print("")
    print("")
    
    global counternum
    global starttime
    
    
    while True:
        if counternum == 1:
            for i in range (3000):
                sys.stdout.write('\033[1A') # Up a line
                sys.stdout.write('\033[1A') # Up a line
                sys.stdout.write(' \r\033[K') # Delete current line
                print('Сортировка файлов: {0} из 200000'.format(i))
                print('Затраченное время: {0}'.format(i))
                sys.stdout.flush()
            counternum = 2
        else:
            for i in range (2000):
                sys.stdout.write('\033[1A') # Up a line
                sys.stdout.write('\033[1A') # Up a line
                sys.stdout.write(' \r\033[K') # Delete current line
                print('Перемещение файлов: {0} из 200000'.format(i))
                print('Затраченное время: {0}'.format(i))
                sys.stdout.flush()
            counternum == 1
            

            

def PrintToScreen(text):
    global counternum
    global starttime
    
    result = time.time() - starttime
    result = datetime2.timedelta(seconds=round(result))

    sys.stdout.write(' \r\033[K')
    sys.stdout.write('\033[1A') # Up a line
    sys.stdout.write(' \r\033[K')
    sys.stdout.write('\033[1A') # Up a line
    sys.stdout.write(' \r\033[K') # Delete current line
    print(text)
    print('Затраченное время: {0}'.format(result))
    sys.stdout.flush()
    #sys.stdout.flush()

            

def log(string, color, font="slant", figlet=False):
#log("P D F  M i n e r", color="cyan", figlet=True)
#log("Обработка файла 6 из 15: 383.pdf", "green")

    if colored:
        if not figlet:
            six.print_(colored(string, color))
        else:
            six.print_(colored(pyfiglet.figlet_format(
                string, font=font), color))
    else:
        six.print_(string)

if __name__ == "__main__":
    main()
