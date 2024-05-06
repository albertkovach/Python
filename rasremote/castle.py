import os, sys, time
import pystray
from tkinter import *
from PIL import Image, ImageTk
import keyboard
import paramiko
from pathlib import Path
 


def main():
    global icon

    datafile = "icon.ico"
    if not hasattr(sys, "frozen"):
        datafile = os.path.join(os.path.dirname(__file__), datafile)
    else:
        datafile = os.path.join(sys.prefix, datafile)
    image = Image.open(datafile)

    keyboard.add_hotkey('shift+f1', RunMain) 

    icon = pystray.Icon("Замок", image, "Замок", menu=pystray.Menu (
        pystray.MenuItem("Запуск!", after_click),
        pystray.MenuItem("Выход", after_click)))

    icon.run()


 
def after_click(icon, query):
    if str(query) == "Запуск!":
        RunMain()
    elif str(query) == "Выход":
        icon.stop()



def RunMain():
    OpenSplashScreen()
    PurgeClients()
    KillCurrent()




def OpenSplashScreen():
    icon.stop()
    
    root = Tk()
    root.title("Замок")
    root.configure(background="white")
    root.geometry("514x514+100+100")
    
    datafile = "castle.png"
    if not hasattr(sys, "frozen"):
        datafile = os.path.join(os.path.dirname(__file__), datafile)
    else:
        datafile = os.path.join(sys.prefix, datafile)

    imagefile = Image.open("C:\\Users\\user\\Documents\\GitHub\\Python\\rasremote\\castle.png")
    pilimage = ImageTk.PhotoImage(imagefile)
    img = Label(root, image=pilimage)
    img.image = pilimage
    img.place(x=1, y=1, width=512, height=512)


def PurgeClients():

    array_ip = []
    array_login = []
    array_password = []
    array_port = []
    
    # array_ip.append('192.168.30.54') # Marina
    # array_login.append('home')
    # array_password.append('12345')
    # array_port.append(22)
    
    array_ip.append('192.168.30.51') # Albert
    array_login.append('makel25')
    array_password.append('Linda1294')
    array_port.append(22)
    
    # array_ip.append('192.168.30.38') # Yana
    # array_login.append('яна')
    # array_password.append('2525')
    # array_port.append(22)
    
    
    text_file_path = Path('Z:\log.txt')
    text_file = open(text_file_path, "w")
    text_file = open(text_file_path, "a")


    for i in range(len(array_ip)):
        print('********************************************')
        print('***** Clear {0} of {1} from: {2} \n'.format(i, len(array_ip)-1, array_ip[i]))
        
        try:
            print('** Opening connection...')
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(hostname=array_ip[i], username=array_login[i], password=array_password[i], port=array_port[i])


            print('** Taskkill...')
            try:
                remote_command = "taskkill /IM castle.exe /f"
                print("== Killing castle process DONE")
                stdin, stdout, stderr = client.exec_command(remote_command)
            except Exception as error:
                print('Exception while taskkill castle occurred: {0} \n {1} \n'.format(type(error).__name__, error))
            try:
                remote_command = "taskkill /IM 1cv8c.exe /f"
                stdin, stdout, stderr = client.exec_command(remote_command)
                print("== Killing 1cv8c process DONE")
            except Exception as error:
                print('Exception while taskkill 1cv8c occurred: {0} \n {1} \n'.format(type(error).__name__, error))


            print('** Remove program...')
            try:   
                remote_command = "del C:\castle.exe"
                stdin, stdout, stderr = client.exec_command(remote_command)
                print("== Remove program DONE")
            except Exception as error:
                print('Exception while remove program occurred: {0} \n {1} \n'.format(type(error).__name__, error))


            print('** Remove batch...')
            try:
                remote_command = "del C:\castle.bat"
                stdin, stdout, stderr = client.exec_command(remote_command)
                print("== Remove batch DONE")
            except Exception as error:
                print('Exception while removing batch occurred: {0} \n {1} \n'.format(type(error).__name__, error))


            client.close()
            print('SSH connection to {0} closed \n'.format(array_ip[i]))
            print('******************************************** \n')


        except Exception as error:
            print('Exception while connecting occurred: {0} \n {1} \n'.format(type(error).__name__, error))


def KillCurrent():
    os.system('taskkill /IM 1cv8c.exe /f')
    os.system('c:\\castle.bat')
    
    
    

def resource_path(relative_path):    
    try:       
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)




main()



# SSH execution with output
# remote_command = "dir"
# stdin, stdout, stderr = client.exec_command(remote_command)
# response = stdout.read() 
# decodedresponse = str(response,"cp866")
# error = stderr.read() 
# decodederror = str(error,"cp866")
# lines = decodedresponse.split('\n')
# for line in lines:
    # text_file.write(line)
# if len(decodederror) > 0:
    # text_file.write(decodederror)
