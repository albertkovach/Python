from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
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
    root.geometry('400x300+{}+{}'.format(scrnw, scrnh))
        
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
        global TextLog
        TextLog = Text(background="white", font=("Arial", 9))
        TextLog.place(x=20, y=55, width=350, height=200)
        
        global Btn1
        Btn1 = Button(text='main', command=BtnCmd1, font=("Arial", 11))
        Btn1.place(x=20, y=20)
        
        global Btn2
        Btn2 = Button(text='fake', command=BtnCmd2, font=("Arial", 11))
        Btn2.place(x=100, y=20)


def BtnCmd1():
    Selector("ToMain")
    
def BtnCmd2():
    Selector("ToFake")

def Selector(OperatingMode):
    print(OperatingMode)
    print("*********")
    
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
    elif OperatingMode == "ToFake":
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



if __name__ == '__main__':
    main()




