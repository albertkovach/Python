from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import os, sys
from pathlib import Path
import paramiko

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
    root.geometry('255x130+{}+{}'.format(scrnw, scrnh))

    app = GUI(root)
    root.mainloop()

    
    

class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.parent.title("Deployer")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        # = List of clients
        # =================
        global array_ip
        global array_login
        global array_password
        global array_port
        
        array_ip = []
        array_login = []
        array_password = []
        array_port = []
        
        array_ip.append('192.168.30.54') # Marina
        array_login.append('home')
        array_password.append('12345')
        array_port.append(22)
        
        array_ip.append('192.168.30.51') # Albert
        array_login.append('makel25')
        array_password.append('Linda1294')
        array_port.append(22)
        
        # array_ip.append('192.168.30.38') # Yana
        # array_login.append('яна')
        # array_password.append('2525')
        # array_port.append(22)



        # = Files path
        # =================
        global local_file_path
        global remote_file_path
        
        global local_batch_path
        global remote_batch_path
        
        local_file_path = "C:\\Users\\user\\Documents\\GitHub\\Python\\rasremote\\castle.exe"
        remote_file_path = "C:\\castle.exe"
        
        local_batch_path = "C:\\Users\\user\\Documents\\GitHub\\Python\\rasremote\\castle.bat"
        remote_batch_path = "C:\\castle.bat"



        # = Server data
        # =================
        global host
        global user
        global secret
        global port
        
        global target_infobase_name
        global target_infobase_id

        global main_db_server
        global main_db_user
        global main_db_pwd
        global main_db_name

        global secondary_db_server
        global secondary_db_user
        global secondary_db_pwd
        global secondary_db_name

        host = '192.168.30.252'
        user = 'shooter'
        secret = 'Lbfvtnh670!'
        port = 22
        
        target_infobase_name = "color"
        target_infobase_id = ""

        # vm-resr main
        main_db_server = "192.168.40.12"
        main_db_user = "sa"
        main_db_pwd = "555RRReee$$$"
        main_db_name = "color-r"

        # resr second
        secondary_db_server = "192.168.40.2"
        secondary_db_user = "sa"
        secondary_db_pwd = "666TTTrrr%%%"
        secondary_db_name = "color-b"



        # = User interface
        # =================
        global DeployBtn
        DeployBtn = Button(text='Deploy', command=DeployApps, font=("Arial", 11))
        DeployBtn.place(x=30, y=15, width=90, height=27)
        
        global ClearBtn
        ClearBtn = Button(text='Clear', command=ClearApps, font=("Arial", 11))
        ClearBtn.place(x=130, y=15, width=90, height=27)
        
        global EnableBtn
        EnableBtn = Button(text='Run', command=EnableBase, font=("Arial", 11))
        EnableBtn.place(x=30, y=80, width=90, height=27)
        
        global DisableBtn
        DisableBtn = Button(text='Stop', command=StopBase, font=("Arial", 11))
        DisableBtn.place(x=130, y=80, width=90, height=27)



def DeployApps():
    ConsoleDefName('DeployApps')
    
    for i in range(len(array_ip)):
        print('============================================')
        print('***** Deploying {0} of {1} to: {2}'.format(i, len(array_ip), array_ip[i]))
 
 
        print("** Start connection")
        try:
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(hostname=array_ip[i], username=array_login[i], password=array_password[i], port=array_port[i])


            print("** Running SFTP")
            try:
                sftp = client.open_sftp()
                sftp.put(local_file_path, remote_file_path)
                sftp.put(local_batch_path, remote_batch_path)
                sftp.close()
                print("== Running SFTP DONE")
            except Exception as error:
                print("Exception while SFTP: ", type(error).__name__)
                print(error)
        
        
            print("** Create autorun registry entry")
            try:
                remote_command = 'REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /V "castle" /t REG_SZ /F /D "C:\castle.exe"'
                stdin, stdout, stderr = client.exec_command(remote_command)
                print("== Create autorun registry entry DONE")
            except Exception as error:
                print("Exception while registry: ", type(error).__name__)
                print(error)


        except Exception as error:
            print("Exception while connecting: ", type(error).__name__)
            print(error)
        
        
        client.close()
        print('***** SSH connection to {0} closed'.format(array_ip[i]))
        print('============================================')



def ClearApps():
    ConsoleDefName('ClearApps')
    
    for i in range(len(array_ip)):
        print('********************************************')
        print('***** Clear {0} of {1} from: {2}'.format(i, len(array_ip), array_ip[i]))
    
    
        print("** Start connection")
        try:
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(hostname=array_ip[i], username=array_login[i], password=array_password[i], port=array_port[i])


            print("** Killing process")
            try:
                remote_command = "taskkill /IM castle.exe /f"
                stdin, stdout, stderr = client.exec_command(remote_command)
                print("== Killing process DONE")
            except Exception as error:
                print("Exception while stopping: ", type(error).__name__)
                print(error)


            print("** Deleting program")
            try:
                remote_command = "del C:\castle.exe"
                stdin, stdout, stderr = client.exec_command(remote_command)
                print("== Deleting program DONE")
            except Exception as error:
                print("Exception while deleting: ", type(error).__name__)
                print(error)


            print("** Deleting batch")
            try:
                remote_command = "del C:\castle.bat"
                stdin, stdout, stderr = client.exec_command(remote_command)
                print("== Deleting batch DONE")
            except Exception as error:
                print("Exception while deleting: ", type(error).__name__)
                print(error)

       
            print("** Deleting registry entry")
            try:
                remote_command = 'REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\castle"'
                stdin, stdout, stderr = client.exec_command(remote_command)
                print("== Deleting registry entry DONE")
            except Exception as error:
                print("Exception while deleting: ", type(error).__name__)
                print(error)


            print("** Closing SSH connection...")
            client.close()
            print('***** SSH connection to {0} closed'.format(array_ip[i]))
            print('********************************************')
        
        
        except Exception as error:
            print("Exception while connecting:", type(error).__name__)
            print(error)



def EnableBase():
    ConsoleDefName('EnableBase')

    CloseClients()
    Switcher('ToMain')



def StopBase():
    ConsoleDefName('StopBase')

    CloseClients()
    Switcher('ToSecond')



def CloseClients():
    for i in range(len(array_ip)):
        print('***** Closing {0} of {1} from: {2}'.format(i+1, len(array_ip)-1, array_ip[i]))
        
        try:
            print('** Opening connection...')
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(hostname=array_ip[i], username=array_login[i], password=array_password[i], port=array_port[i])


            print('** Taskkill 1cv8c ...')
            try:
                remote_command = "taskkill /IM 1cv8c.exe /f"
                stdin, stdout, stderr = client.exec_command(remote_command)
                print("== Taskkill 1cv8c DONE")
            except Exception as error:
                print('Exception while taskkill 1cv8c occurred: {0} \n {1} \n'.format(type(error).__name__, error))


        except Exception as error:
            print('Exception while connecting occurred: {0} \n {1} \n'.format(type(error).__name__, error))



def Switcher(OperatingMode):
    print('')
    print('')
    print('*************************')
    print("++  Infobase operation ++")
    print("*************************")
    print('+++++ ' + OperatingMode + ' +++++')
    print("*************************")
    
    
    host = '192.168.30.252'
    user = 'shooter'
    secret = 'Lbfvtnh670!'
    port = 22

    cluster_id = ""
    cluster_infobases_names = []
    cluster_infobases_id = []
    infobase_sessions_id = []
    
    target_infobase_name = "color"
    target_infobase_id = ""

    # vm-resr - red - main
    main_db_server = "192.168.40.12"
    main_db_user = "sa"
    main_db_pwd = "555RRReee$$$"
    main_db_name = "color-r"

    # resr - blue - fake
    secondary_db_server = "192.168.40.2"
    secondary_db_user = "sa"
    secondary_db_pwd = "666TTTrrr%%%"
    secondary_db_name = "color-b"


    # *** Open connection
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(hostname=host, username=user, password=secret, port=port)


    # *** Find cluster id 
    parse_cluster_command = "rac cluster list"
    stdin, stdout, stderr = client.exec_command(parse_cluster_command)
    response = stdout.read() 
    decodedresponse = str(response,"cp866")
    error = stderr.read() 
    decodederror = str(error,"cp866")
    lines = decodedresponse.split('\n')
    for line in lines:
        if line.startswith('cluster'):
            line = line.replace(" ", "")
            cluster_index = line.find(":")
            if cluster_index != -1:
                cluster_id = line[cluster_index + 1:]
            else:
                print("error parsing cluster id")
    if len(decodederror) > 0:
        print(decodederror)
    print("cluster_id " + cluster_id)


    # *** Find target infobase id 
    parse_infobaselist_command = "rac infobase --cluster=" + cluster_id + " --cluster-user=shooter --cluster-pwd=Lbfvtnh670! summary list"
    stdin, stdout, stderr = client.exec_command(parse_infobaselist_command)
    response = stdout.read() 
    decodedresponse = str(response,"cp866")
    error = stderr.read() 
    decodederror = str(error,"cp866")
    lines = decodedresponse.split('\n')
    for line in lines:
        if line.startswith('infobase'):
            line = line.replace(" ", "")
            subline_index = line.find(":")
            if subline_index != -1:
                cluster_infobases_id.append(line[subline_index + 1:])
        elif line.startswith('name'):
            line = line.replace(" ", "")
            subline_index = line.find(":")
            if subline_index != -1:
                cluster_infobases_names.append(line[subline_index + 1:])
    error = stderr.read() 
    decodederror = str(error,"cp866")
    if len(decodederror) > 0:
        print(decodederror)

    for i in range(len(cluster_infobases_names)):
        infobase = str(cluster_infobases_names[i])
        infobase = infobase.replace(" ", "")
        infobase = infobase.replace(chr(13), "")
        
        if infobase == target_infobase_name:
            target_infobase_id = cluster_infobases_id[i]
            break
    print("target_infobase_id " + target_infobase_id)


    # *** Lock infobase
    lock_infobase_command = "rac infobase --cluster=" + cluster_id + " --cluster-user=shooter --cluster-pwd=Lbfvtnh670! update --infobase=" + \
                            target_infobase_id + " --sessions-deny=on --scheduled-jobs-deny=on --license-distribution=deny"
    stdin, stdout, stderr = client.exec_command(lock_infobase_command)
    response = stdout.read() 
    decodedresponse = str(response,"cp866")
    error = stderr.read() 
    decodederror = str(error,"cp866")
    if len(decodedresponse) > 0:
        print(decodedresponse)
    if len(decodederror) > 0:
        print(decodederror)
    print("infobase locked")



    # *** Clear infobase sessions
    get_infobase_sessions_command = "rac session --cluster=" + cluster_id + " --cluster-user=shooter --cluster-pwd=Lbfvtnh670! list --infobase=" + \
                            target_infobase_id
    stdin, stdout, stderr = client.exec_command(get_infobase_sessions_command)
    response = stdout.read() 
    decodedresponse = str(response,"cp866")
    error = stderr.read() 
    decodederror = str(error,"cp866")
    if len(decodederror) > 0:
        print(decodederror)
    lines = decodedresponse.split('\n')
    for line in lines:
        if line.startswith('session '):
            line = line.replace(" ", "")
            subline_index = line.find(":")
            if subline_index != -1:
                infobase_sessions_id.append(line[subline_index + 1:])

    for session in infobase_sessions_id:
        print("terminating session " + session)
        terminate_session_command = "rac session --cluster=" + cluster_id + " --cluster-user=shooter --cluster-pwd=Lbfvtnh670! terminate --session=" + session
        stdin, stdout, stderr = client.exec_command(terminate_session_command)
        response = stdout.read() 
        decodedresponse = str(response,"cp866")
        error = stderr.read() 
        decodederror = str(error,"cp866")
        if len(decodedresponse) > 0:
            print(decodedresponse)
        if len(decodederror) > 0:
            print(decodederror)


    # *** Replace infobase
    if OperatingMode == "ToMain":
        replace_infobase_command = "rac infobase --cluster=" + cluster_id + " --cluster-user=shooter --cluster-pwd=Lbfvtnh670! update --infobase=" + \
                                target_infobase_id + " --db-server=" + main_db_server + \
                                " --db-name=" + main_db_name + " --db-user=" + main_db_user + " --db-pwd=" + main_db_pwd
    elif OperatingMode == "ToSecond":
        replace_infobase_command = "rac infobase --cluster=" + cluster_id + " --cluster-user=shooter --cluster-pwd=Lbfvtnh670! update --infobase=" + \
                                target_infobase_id + " --db-server=" + secondary_db_server + \
                                " --db-name=" + secondary_db_name + " --db-user=" + secondary_db_user + " --db-pwd=" + secondary_db_pwd
    stdin, stdout, stderr = client.exec_command(replace_infobase_command)
    response = stdout.read() 
    decodedresponse = str(response,"cp866")
    error = stderr.read() 
    decodederror = str(error,"cp866")
    if len(decodedresponse) > 0:
        print(decodedresponse)
    if len(decodederror) > 0:
        print(decodederror)
    print("infobase replaced")


    # *** Unlock infobase
    unlock_infobase_command = "rac infobase --cluster=" + cluster_id + " --cluster-user=shooter --cluster-pwd=Lbfvtnh670! update --infobase=" + \
                            target_infobase_id + " --sessions-deny=off --scheduled-jobs-deny=off --license-distribution=allow"
    stdin, stdout, stderr = client.exec_command(unlock_infobase_command)
    response = stdout.read() 
    decodedresponse = str(response,"cp866")
    error = stderr.read() 
    decodederror = str(error,"cp866")
    if len(decodedresponse) > 0:
        print(decodedresponse)
    if len(decodederror) > 0:
        print(decodederror)
    print("infobase unlocked, done")


    # *** Close connection
    client.close()
    
    print('')
    print("** Infobase operation DONE")
    print("**************************")



def ConsoleDefName(DefName):
    print('')
    print('============================================')
    ConsoleLine = '   {0}   '.format(DefName)
    
    direction = True
    finished = False
    while not finished:
        if len(ConsoleLine) == 43:
            finished = True
            
        if direction:
            ConsoleLine = '=' + ConsoleLine
            direction = False
        else:
            ConsoleLine = ConsoleLine + '='
            direction = True
        
    print(ConsoleLine)
    print('')
    print('')  



if __name__ == '__main__':
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