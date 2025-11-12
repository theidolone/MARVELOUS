from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import os, sys, random
import simpleaudio as sa

user = os.path.expanduser("~")      #grab the current user: c:/users/(name)/
startup = "Appdata\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup"        #defining the path to the startup folder
fullpath = os.path.join(user, startup)      #joining the user and the startup path to form a full absolute path

if getattr(sys, 'frozen', False):
    cwd = os.path.dirname(sys.executable)

else:
    cwd = os.path.dirname(os.path.abspath(__file__))        

def resource_path(relative_path):

    if getattr(sys, 'frozen', False):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))

    else:

        base_path = cwd

    return os.path.abspath(os.path.join(base_path, relative_path))

try: 
    marvelous = open(f"{fullpath}\\marvelous.vbs", "r+")     #joining the absolute path to startup with the name of the vbs script file to see if it exists
    marvelous.close()        

except FileNotFoundError:
    marvelous = open(f'{fullpath}\\marvelous.vbs', 'w')      #creating the taiga.vbs file if it doesn't exist
    marvelous.write(f'Set WshShell = CreateObject("WScript.Shell")\nWshShell.Run """{cwd}\\marvelous.exe""", 0, False')      #writing instructions for vbs to execute
    marvelous.close()

today = datetime.now()
sundaycheck = today.strftime("%A")

if sundaycheck == "Sunday":
    sundaypath = resource_path("marvelous.wav")
    wave_obj = sa.WaveObject.from_wave_file(sundaypath)
    play_obj = wave_obj.play()


    root = tk.Tk()      
    root.withdraw()
    root.attributes("-topmost", True)
    root.iconbitmap(resource_path("marvelous.ico"))
    tk.messagebox.showinfo(title="MARVELOUS!!", message="It's Marvelous Sunday Time!")

    root.destroy()
    play_obj.stop()
    sys.exit()

else:
    sundayRoll = random.randint(1,4)
    if sundayRoll == 4:
        sundaypath = resource_path("marvelous.wav")
        wave_obj = sa.WaveObject.from_wave_file(sundaypath)
        play_obj = wave_obj.play()


        root = tk.Tk()      
        root.withdraw()
        root.attributes("-topmost", True)
        root.iconbitmap(resource_path("marvelous.ico"))
        tk.messagebox.showinfo(title=f"MARVELOUS {sundaycheck}!?", message=f"It's Marvelous {sundaycheck} Time!")

        root.destroy()
        play_obj.stop()
        sys.exit()

    else:
        sys.exit()