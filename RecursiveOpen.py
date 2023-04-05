import subprocess
import win32com.client 
import time
import os
import psutil
import glob
import webbrowser 

def executeLocal(executable):
    subprocess.Popen([executable])

def executeWeb(url):
    webbrowser.open(url)  
    
def memOcupate():
    cpuUsed = psutil.cpu_percent(5)
    ramUsed = psutil.virtual_memory()[2]
    if cpuUsed > 80 or ramUsed > 80:
        return True
    else:
        return False

def main():
    eof = False
    path = os.getcwd()
    print("Current path: " + path)
    while not eof:
        eof = True
        localFiles = glob.glob(path + "\*")
        for value in localFiles:
            file = os.path.splitext(value)

            if file[1] == ".lnk":
                while memOcupate():
                    time.sleep(5)
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(value)
                executeLocal(shortcut.Targetpath)
            elif file[1] == ".url":
                while memOcupate():
                    time.sleep(5)
                executeWeb(value)  
            elif file[1] == "":
                path = value
                eof = False
                            
main()