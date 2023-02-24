import sys
from PyQt5.QtWidgets import QApplication,QWidget        # 위젯관련 함수 사용하기
import pyautogui                          # pyautogui 함수 사용하기
import time
import tkinter
from tkinter import ttk
from tkinter import filedialog,messagebox


from threading import Thread
import json

class MacroData():
    
    def __init__(self,xPos,yPos,sec,name):
        self.xPos = xPos
        self.yPos = yPos
        self.sec = sec
        self.name = name
    def toString(self):
        return f'{self.name} : xPos = {self.xPos}, yPos = {self.yPos}, sec = {self.sec}'

class MyApp(QWidget):

    def __init__(self):
        super().__init__()                # 부모 클래스 초기화
        
        # pyautogui.mouseInfo()
        self.list = []
        
        self.root = tkinter.Tk()
        self.root.title('클릭 매크로')
        self.root.minsize(500, 100)
        mainframe = ttk.Frame(self.root, padding='3 3 12 12')
        mainframe.grid(column=0, row=0, sticky=(tkinter.N, tkinter.W, tkinter.E, tkinter.S))

        self.logTextboxSV         = tkinter.StringVar() # The str contents of the log text area.

        CUR_ROW = 1
        ttk.Label(mainframe, text='이름').grid(column=1, row=CUR_ROW, sticky=tkinter.W)
        self.nameTextboxSV = tkinter.StringVar() # The str contents of the xy text field.
        self.nameTextbox = ttk.Entry(mainframe,  textvariable=self.nameTextboxSV)
        self.nameTextbox.grid(column=2,columnspan=6, row=CUR_ROW, sticky=(tkinter.W, tkinter.E))
        
        CUR_ROW+=1
        ttk.Label(mainframe, text='X,Y Pos ').grid(column=1, row=CUR_ROW, sticky=tkinter.W)
        self.xTextboxSV = tkinter.StringVar() # The str contents of the xy text field.
        self.xInfoTextbox = ttk.Entry(mainframe, width=10, textvariable=self.xTextboxSV,state='readonly')
        self.xInfoTextbox.grid(column=2,row=CUR_ROW, sticky=(tkinter.W, tkinter.E))
        
        self.yTextboxSV = tkinter.StringVar() # The str contents of the xy text field.
        self.yInfoTextbox = ttk.Entry(mainframe, width=10, textvariable=self.yTextboxSV,state='readonly')
        self.yInfoTextbox.grid(column=3, row=CUR_ROW, sticky=(tkinter.W, tkinter.E))


        ttk.Label(mainframe, text='대기 시간').grid(column=4, row=CUR_ROW, sticky=tkinter.W)
        self.SecTextboxSV = tkinter.IntVar() # The str contents of the xy text field.
        self.SecInfoTextbox = ttk.Entry(mainframe, width=10, textvariable=self.SecTextboxSV)
        self.SecInfoTextbox.grid(column=5, row=CUR_ROW, sticky=(tkinter.W, tkinter.E))
        
        
        self.SecUpSV = tkinter.StringVar()
        self.SecUpSV.set('▲')
        self.SecUpButton = ttk.Button(mainframe,width=5, textvariable=self.SecUpSV, command=self._upSec)
        self.SecUpButton.grid(column=6, row=CUR_ROW, sticky=tkinter.W)
        self.SecUpButton.bind('<Return>', self._upSec)
        
        self.SecDownSV = tkinter.StringVar()
        self.SecDownSV.set('▼')
        self.SecDownButton = ttk.Button(mainframe,width=5, textvariable=self.SecDownSV,  command=self._downSec)
        self.SecDownButton.grid(column=7, row=CUR_ROW, sticky=tkinter.W)
        self.SecDownButton.bind('<Return>', self._downSec)
        
        self.addListButtonSV = tkinter.StringVar()
        self.addListButtonSV.set('추가 (F1)')
        self.addListButton = ttk.Button(mainframe,width=10, textvariable=self.addListButtonSV,  command=self._addListBox)
        self.addListButton.grid(column=9,row=CUR_ROW, sticky=tkinter.W)
        self.addListButton.bind('<Return>', self._addListBox)
        

        
        CUR_ROW+=1
        self.listbox = tkinter.Listbox(mainframe,height=15,selectmode='single')
        self.listbox.grid(column=1, row=CUR_ROW, columnspan=8, sticky=(tkinter.W, tkinter.E, tkinter.N, tkinter.S))
        self.listboxScrollbar = ttk.Scrollbar(mainframe, orient=tkinter.VERTICAL, command=self.listbox.yview)
        self.listboxScrollbar.grid(column=8, row=CUR_ROW, sticky=(tkinter.N, tkinter.S))
        self.listbox['yscrollcommand'] = self.listboxScrollbar.set

        self.deleteListButtonSV = tkinter.StringVar()
        self.deleteListButtonSV.set('삭제')
        self.deleteListButton = ttk.Button(mainframe, textvariable=self.deleteListButtonSV, command=self._deleteListBox)
        self.deleteListButton.grid(column=9, row=CUR_ROW, sticky=(tkinter.W, tkinter.E, tkinter.N, tkinter.S))
        self.deleteListButton.bind('<Return>', self._deleteListBox)


        CUR_ROW+=1
        ttk.Label(mainframe, text='실행 횟수').grid(column=1, row=CUR_ROW, sticky=tkinter.W)
        self.CountTextboxSV = tkinter.StringVar() # The str contents of the xy text field.
        self.CountTextbox = ttk.Entry(mainframe, width=10,textvariable=self.CountTextboxSV,state='readonly')
        self.CountTextbox.grid(column=2, row=CUR_ROW, sticky=(tkinter.W, tkinter.E))
        


        # CUR_ROW+=1
        self.StartBottnSV = tkinter.StringVar()
        self.StartBottnSV.set('실행 (F2)')
        self.StartButton = ttk.Button(mainframe, textvariable=self.StartBottnSV,  command=self._start)
        self.StartButton.grid(column=3, row=CUR_ROW, sticky=tkinter.W)
        self.StartButton.bind('<Return>', self._start)
        
        # CUR_ROW+=1
        self.StopBottnSV = tkinter.StringVar()
        self.StopBottnSV.set('중지 (F3)')
        self.StopButton = ttk.Button(mainframe, textvariable=self.StopBottnSV,  command=self._stop)
        self.StopButton.grid(column=4, row=CUR_ROW,  sticky=tkinter.W)
        self.StopButton.bind('<Return>', self._stop)

        # CUR_ROW+=1
        self.ExitBottnSV = tkinter.StringVar()
        self.ExitBottnSV.set('종료 (F4)')
        self.ExitButton = ttk.Button(mainframe, textvariable=self.ExitBottnSV,command=self._exit)
        self.ExitButton.grid(column=5, row=CUR_ROW, columnspan=8, sticky=tkinter.W)
        self.ExitButton.bind('<Return>', self._exit)
        
        
        self.root.option_add('*tearOff', tkinter.FALSE) # Disable tkinter's ugly tear-off menus which are enabled by default.

        menu = tkinter.Menu(self.root)
        self.root.config(menu=menu)

        copyMenu = tkinter.Menu(menu)
        copyMenu.add_command(label='save', command=self._saveAs, accelerator='F5', underline=5)
        copyMenu.add_command(label='load', command=self._load, accelerator='F6', underline=5)
        menu.add_cascade(label='Option', menu=copyMenu, underline=0)

   
        self.root.bind_all('<F1>', self._addListBox)
        self.root.bind_all('<F2>', self._start)
        self.root.bind_all('<F3>', self._stop)
        self.root.bind_all('<F4>', self._exit)
        self.root.bind_all('<F5>', self._saveAs)
        self.root.bind_all('<F6>', self._load)


        
        self._init()
        
        self.root.resizable(False, False) # Prevent the window from being resized.

        self.root.focus() # Put the focus on the XY coordinate text field to start.
        
        # self.root.attributes('-topmost', True)
        self.root.protocol('WM_DELETE_WINDOW',self._exit)
        self.root.mainloop()


    def _load(self, *args):
        filepath=filedialog.askopenfile()
        if filepath.name != None:
            self.list.clear()
            self.listbox.delete(0,'end')
            tmp=''
            f=open(filepath.name,'r')
            while True:
                c = f.read()
                if c == '':
                    break
                tmp  = c + tmp
            f.close()
            loadData = json.loads(tmp)
            for item in loadData:
                data = MacroData(item['xPos'],item['yPos'],item['sec'],item['name'])
                self.list.append(data)
                self.listbox.insert(self.listbox.size(),data.toString())

    def _saveAs(self, *args):
        filepath = filedialog.asksaveasfile()
        if filepath.name != None:
            with open(filepath.name, "w",encoding="UTF-8") as file:
                saveData = json.dumps(self.list, default=lambda o: o.__dict__, sort_keys=True, indent=4)
                file.write(saveData)
            

    def _init(self):
        self.xTextboxSV.set(1)
        self.yTextboxSV.set(1)
        self.SecTextboxSV.set(1)
        self.CountTextboxSV.set(0)

    def _upSec(self,*args):
        sec = int(self.SecTextboxSV.get())
        sec+=1
        self.SecTextboxSV.set(sec)

    def _downSec(self,*args):
        sec = int(self.SecTextboxSV.get())
        if sec > 0:
            sec-=1
            self.SecTextboxSV.set(sec)

        
    def _addListBox(self,*args):
        try:
            x,y = pyautogui.position()
            self.xTextboxSV.set(str(x))
            self.yTextboxSV.set(str(y))
            data = MacroData(self.xTextboxSV.get(),self.yTextboxSV.get(),self.SecTextboxSV.get(),self.nameTextboxSV.get())
            self.list.append(data)
            self.listbox.insert(self.listbox.size(),data.toString())
            topOfTextArea, bottomOfTextArea = self.listbox.yview()
            self.listbox.yview_moveto(bottomOfTextArea)
            self.nameTextbox.focus()
        except tkinter.TclError:
            messagebox.showerror('경고','입력한 대기시간이 숫자가 아닙니다.')
        # self._init()

    def _deleteListBox(self,*args):
        listIdx = self.listbox.curselection()
        if len(listIdx)==1:
            self.list.pop(listIdx[0])
            self.listbox.delete(listIdx[0])
        
    
    
        
    def _start(self,*args):
        if len(self.list) > 0:
            self._loop=True
            self._stateButton('disable')
            thread = Thread(target=self._macroRun, args=())
            thread.daemon = True
            thread.start()

    def _macroRun(self):
        self.listbox.select_clear(0,self.listbox.size())
        while self._loop :
            for idx,item in enumerate(self.list):
                self.listbox.select_set(idx)
                if self._loop == False:
                    break
                pyautogui.click(int(item.xPos),int(item.yPos))                 # pyautogui의 click 함수 사용 위치는 x,y 좌표
                time.sleep(int(item.sec))
                self.listbox.select_clear(idx)
            self.addCount()
        self.CountTextboxSV.set(0)



    def addCount(self):
        count = int(self.CountTextboxSV.get())
        count+=1
        self.CountTextboxSV.set(count)

    
    def _stop(self,*args):
        self._loop=False
        self._stateButton('enable')
        
    def _stateButton(self,state):
        self.SecUpButton.config(state=state)
        self.SecDownButton.config(state=state)
        self.addListButton.config(state=state)
        self.deleteListButton.config(state=state)
        self.StartButton.config(state=state)
    
    def _exit(self,*args):
        self._loop=False
        self.root.destroy()


if __name__ == '__main__':           
                                     
    app = QApplication(sys.argv)     
    ex = MyApp()                     
