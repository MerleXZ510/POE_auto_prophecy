import win32gui, win32con, win32api
import time
from pynput.keyboard import Key,Listener,Controller
import pyperclip 
import json
import pandas as pd
import threading
import os

class Keyboard(object):
    
    def __init__(self):
        self.a = False
        
    def do_sth(self):
        while True:
            if self.a:
                break
            else:
                print('keep running')
                wish()
        print('end loop')
                
    def press(self, key):
        if key == Key.end: 
            self.a = True
            return False
        else:
            return True
            
    def main(self):
        with Listener(on_press=self.press) as listener:
            listener.join()
            
    def pool(self):
        t1 = threading.Thread(target=self.do_sth, name='LoopThread1')
        t2 = threading.Thread(target=self.main, name='LoopThread2')
        t1.start()
        t2.start()
        t1.join()
        t2.join()

def load_poss():
    with open('poss_data.json') as json_file:  
        data = json.load(json_file)
    poss_navali = data['poss_navali']
    poss_stack = data['poss_stack']
    poss_talk = data['poss_talk']
    return  poss_navali , poss_stack , poss_talk

def load_log():
    with open('logdata.json') as json_file:  
        data = json.load(json_file)
    list_log_name = data['Name']
    list_log_time = data['Time']

    return  list_log_name , list_log_time 

def save_poss(poss_navali,poss_stack,poss_talk):
    poss_data={
        "poss_navali" : poss_navali,
        "poss_stack" : poss_stack,
        "poss_talk" : poss_talk
    }
    with open('poss_data.json', 'w') as outfile:  
        json.dump(poss_data, outfile)
        print('存檔成功!')

def save_log():
    global list_log_name,list_log_time
    log_data={
        "Name" : list_log_name,
        "Time" : list_log_time}
    with open('logdata.json', 'w') as outfile:  
        json.dump(log_data, outfile)
        print('儲存log檔成功!')

    train = pd.DataFrame.from_dict(log_data)
    train.to_excel("logdata.xlsx")

def clickmouseleft(poss,del_time):
    win32api.SetCursorPos(poss)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN|win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(del_time)

def del_item():
    #拿起來
    clickmouseleft(poss_item,delay_time)
    clickmouseleft(poss_del,delay_time)
    clickmouseleft(poss_del,delay_time)

def movePage(pp):
    global page
    keyboard = Controller() 
    if pp > 0:
        for i in range(int(pp)):
            keyboard.press(Key.right) 
            keyboard.release(Key.right)
            time.sleep(0.03)
            page+=1
    elif pp < 0:
        pp *=-1
        for i in range(int(pp)):
            keyboard.press(Key.left)
            keyboard.release(Key.left)
            time.sleep(0.03)
            page -=1

def save_item(i,item):
    global page
    p = list_stack[i] - page
    clickmouseleft(poss_stack,delay_time)
    movePage(p)
    keyboard = Controller()   
    with keyboard.pressed(Key.ctrl): 
        clickmouseleft(poss_item,delay_time) 
    ##確認是否存入   
    win32api.SetCursorPos(poss_item_2) 
    time.sleep(0.2)
    with keyboard.pressed(Key.ctrl): #按下shift
        keyboard.press("c") #shift + a
        time.sleep(0.02)
        keyboard.release("c")
    time.sleep(0.1)
    win32api.SetCursorPos(poss_item)
    time.sleep(0.2)
    with keyboard.pressed(Key.ctrl): #按下shift
        keyboard.press("c") #shift + a
        time.sleep(0.02)
        keyboard.release("c")
    time.sleep(0.2)
    if item in pyperclip.paste():
        save_log()
        os._exit(0)

def check_item(item):
    it = item.replace("\r\n",'').replace("稀有度: 普通",'').replace(" I","")
    list_log_time[list_log_name.index(it)] += 1
    print(list_log_time)
    for i in range(len(list_save)):
        if list_save[i] in item:
            save_item(i,it)
            print("Save")
    if not any(s in item for s in list_save):
        print("Del")
        del_item()


def wish():
    keyboard = Controller() 
    #按空白初始化
    keyboard.press(Key.space)
    keyboard.release(Key.space)
    time.sleep(delay_time) 
    #點老奶奶
    clickmouseleft(poss_navali,delay_time)
    #尋求預言
    clickmouseleft(poss_talk,delay_time)
    #封印
    time.sleep(0.1)
    clickmouseleft(poss_seal,delay_time)
    #封印
    clickmouseleft(poss_seal_yes,delay_time)
    #放下物品
    clickmouseleft(poss_item,delay_time)
    #刪除或者保留
    win32api.SetCursorPos(poss_item)
    with keyboard.pressed(Key.ctrl): #按下shift
        keyboard.press("c") #shift + a
        time.sleep(0.02)
        keyboard.release("c")
    time.sleep(0.02) 
    item = pyperclip.paste().split("--------")[0]   
    print(item)
    check_item(item)



def on_press(key): 
    global poss_navali, poss_stack, poss_talk
    if key == Key.f2:
        x,y=win32api.GetCursorPos()
        poss_navali = (x,y)
        print('設定老奶奶位置:',poss_navali)
        save_poss(poss_navali,poss_stack,poss_talk)
    if key == Key.f3:
        x,y=win32api.GetCursorPos()
        poss_stack = (x,y)
        print('設定倉庫位置:',poss_stack)
        save_poss(poss_navali,poss_stack,poss_talk)
    if key == Key.f4:
        x,y=win32api.GetCursorPos()
        poss_talk = (x,y)
        print('設定尋求預言位置:',poss_talk)
        save_poss(poss_navali,poss_stack,poss_talk)


def on_release(key):
    if key == Key.home:
        print("開始洗")
        return False




with open('delaytime.json') as json_file:  
        data = json.load(json_file)
delay_time = data['delaytime']
t = 0
poss_navali, poss_stack, poss_talk = load_poss()
#封印
poss_seal = (351,520)
#封印-是
poss_seal_yes = (704,478) 
#刪除
poss_del = (934,474)
#物品位置
poss_item = (1093,522)
poss_item_2 = (1137,526)
#---------------------------------------讀取資料
df = pd.read_excel("seal.xlsx")
df["要"] > 0
df_save = df[df["要"] > 0]
list_save = df_save['預言名稱'].values
list_stack = df_save['要'].values
print(list_save)
print(list_stack)
#---------------------------------------
page = 1
df_all = df[df["要"] > -1]
if os.path.isfile('logdata.json'):
    list_log_name, list_log_time = load_log()    
else:
    list_log_name = list(df_all['預言名稱'].values)
    list_log_time = [ 0 for i in list_log_name]   


hwnd = win32gui.FindWindow(None,"Path of Exile")
win32gui.MoveWindow(hwnd,0,0,1600,900,True)
#---------------------------------------




with Listener(on_press=on_press,on_release=on_release) as listener:
    listener.join()
    
#點老奶奶
keyboard = Controller()
clickmouseleft(poss_navali,delay_time)
keyboard.press(Key.space)
keyboard.release(Key.space)
time.sleep(delay_time)
clickmouseleft(poss_stack,delay_time)
movePage(-40)  

page = 1
page = 1

wishloop = Keyboard()
wishloop.pool()

save_log()
