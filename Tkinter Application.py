import tkinter as tk
import os
from win32com.client import Dispatch
import webbrowser
import docx
basic=tk.Tk()
basic.title("Menu driven program")
basic.geometry('500x500')
string1="Menu Driven Program anything you want to do?"
windows_speak=Dispatch('SAPI.Spvoice')
e1=tk.Entry(basic)
mainlabel=tk.Label(basic,text="Enter your text here :",bg="black",fg="white")
def helpb():
    window=tk.Tk()
    window.geometry('320x220')
    stri="WHAT YOU CAN DO\nYou can access:-\nnotepad\memo \n chrome\nWindows media player\npython\ncreate hello world python\ncreate msword\nCMD\nyoutube\ngoogle\nfacebook\nfolder"
    helplab=tk.Label(window,text=stri,fg="yellow",bg="black")
    helplab.place(x=100,y=0)
    window.configure(bg="black")   
another=tk.Button(basic,text="Help",bg="yellow",fg="Red",command=helpb)
e1.place(x=183,y=300)
def action():
    string=e1.get()
    hey=(("open" in string) or ("run"in string) or("execute" in string) or ("perform"in string))
    if((hey==True) and (("notepad" in string)or("memo" in string))):
        windows_speak.speak('opening notepad')
        os.system("notepad")
        windows_speak.speak('Is there anything else you want to do?')
    elif(((hey==True)or("browse" in string) or("search" in string))and (("google chrome"in string) or ("browser"in string)or("internet"in string))):
        windows_speak.speak('opening chrome')
        os.system("chrome")
        windows_speak.speak('Is there anything else you want to do?')
    elif((hey==True)and (("command prompt"in string) or ("cmd"in string)or("prompt"in string))):
        windows_speak.speak('opening command prompt')
        os.system("cmd")
        windows_speak.speak('Is there anything else you want to do?')
    elif((hey==True) and (("wmplayer"in string)or ("media player" in string)or ("player" in string))):
         os.system('wmplayer')
         windows_speak.speak('If windows media player isnt opening set environment variable')
         windows_speak.speak('Is there anything else you want to do?')
    elif((hey==True)and ("youtube"in string) or ("videos" in string) or ("internet" in string)):
        webbrowser.open('https://youtube.com')
        windows_speak.speak("opening youtube")
    elif((hey==True)and("facebook" in string)or ("fb"in string)):
        webbrowser.open('https://facebook.com')
        windows_speak.speak("opening facebook")
        windows_speak.speak('Is there anything else you want to do?')
    elif((hey==True)and(("google"in string))):
        webbrowser.open('https://google.co.in')
        windows_speak.speak("opening google")
    elif((hey==True)and ("python"in string) or ("programming"in string)):
        if((hey==True)and ("simple"in string) or ("python program"in string) or("hello world" in string)):
            document=docx.Document()
            document.add_paragraph('print(''hello world'')')
            document.save('hello_world.py')
            windows_speak.speak("Hello world program is generated in Desktop or search it in taskbar")
        else:
            os.system('python')
            windows_speak.speak("opening python for you")
        
    elif((hey==True)and("word"in string)or("document"in string)or("docx"in string)):
        mydoc=docx.Document()
        mydoc.save('auto_generated_docx')
        windows_speak.speak('automatic generated windows document in Desktop')
        
        
    elif((hey==True)and("directory"in string)or("folder" in string) or("folders"in string) or("desktop" in string)):
        webbrowser.open('')
        windows_speak.speak("Opening Desktop Folders")
    elif(("quit" in string)or("exit"in string)):
        windows_speak.speak('closing the terminal Take care')
        basic.destroy()
lab=tk.Label(basic,text="Turn On Your Spekear\n\nclick on help to see what you can do",fg='black',bg='yellow')
label=tk.Label(basic,text="Work it Like you Talk it",height=2,width=25,fg="yellow",bg="black")
label2=tk.Label(basic,text="Type Anything You want to Do \nor click on help to see what you can do \n or exit terminal by pressing exit",fg="white",bg="black")
la=tk.Label(basic,text="please click on work it button\nTo pass on your command",fg='black',bg='white')
send=tk.Button(basic,text="Work It",height=0,width=5,fg="black",bg="yellow",command=action)
lab.place(x=147,y=0)
la.place(x=164,y=325)
mainlabel.place(x=188,y=279)
another.place(x=225,y=400)
send.place(x=220,y=365)
label2.place(x=137,y=200)
label.place(x=150,y=100)
basic.configure(bg="black")
windows_speak.speak(string1)
basic.mainloop()
