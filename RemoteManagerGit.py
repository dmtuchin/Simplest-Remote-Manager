# https://pypi.org/project/pyTelegramBotAPI/0.3.0/   
#

import os
import wmi
import sys

import time
import telebot
import psutil
import pyautogui

import ctypes
from ctypes import wintypes

import psutil
import platform
import subprocess

GetWindowThreadProcessId = ctypes.windll.user32.GetWindowThreadProcessId
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW

TOKEN = 'Set here YOUR Telegram bot TOKEN'
bot = telebot.TeleBot(TOKEN)

def extract_arg(arg):
    return arg.split()[1:]

def blocks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

def send_multiline(chatid, alist, block_sz = 10):
    chunks = list(blocks(alist, block_sz))
    for ch in chunks:
        multiline = '\n'.join(ch)
        bot.send_message(chatid, multiline)

@bot.message_handler(commands=['help'])
def start_command(message):
    bot.send_message(message.chat.id, "See help on apllication window.")

@bot.message_handler(commands=['aw'])
def aw_command(message): # Windows - hwnd of foreground window
    if platform.system() == 'Windows':
        lpdw_process_id = ctypes.c_ulong()
        hWnd = ctypes.windll.user32.GetForegroundWindow()
        length = GetWindowTextLength(hWnd)
        buf = ctypes.create_unicode_buffer(length + 1)
        ctypes.windll.user32.GetWindowTextW(hWnd, buf, length + 1)
        result = GetWindowThreadProcessId(hWnd, ctypes.byref(lpdw_process_id))
        process_id = lpdw_process_id.value
        bot.send_message(message.chat.id, "The ProcessId of foreground window is " + str(process_id) + "\n" + buf.value)

@bot.message_handler(commands=['k'])
def kill_command(message):
    pid = extract_arg(message.text)
    if len(pid) > 0:
        prc = psutil.Process(int(pid[0]))
        prc.kill()
    
@bot.message_handler(commands=['f'])
def foreground_command(message):
    pid = extract_arg(message.text)
    if len(pid) > 0:
        from pywinauto import Application
        app = Application().connect(process=int(pid[0]))
        app.top_window().set_focus()

@bot.message_handler(commands=['r'])
def foreground_command(message):
    name = extract_arg(message.text)
    if len(name) > 0:
        found = []; path = []
        what = str(name[0]).lower()
        bot.send_message(message.chat.id, "Long time operation. Please wait...")
        # find module on disk
        for root, dirs, files in os.walk("/"):
            for file in files:
                file = file.lower()
                if file.startswith(what):
                    path.append(os.path.join(root, file))
                    found.append(file)#os.path.join(root, file))

        # decision
        if len(found) > 1:
            # find exact
            if found.count(what) == 1:
                ind = found.index(what)
                # start process
                bot.send_message(message.chat.id, path[ind])
                proc = subprocess.Popen(path[ind])
            else:
                send_multiline(message.chat.id, found, block_sz = 50)
                bot.send_message(message.chat.id, "Enter the exact name again")
        elif len(found) == 1:
            # start process
            bot.send_message(message.chat.id, path[0])
            proc = subprocess.Popen(path[0])
        else:
            bot.send_message(message.chat.id, "Found nothing")

@bot.message_handler(commands=['scr'])
def screenshot_command(message):
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(r'fullscreen.png')
    img = open(r'fullscreen.png', 'rb')
    bot.send_document(message.chat.id, img) 

@bot.message_handler(commands=['p'])
def process_command(message):
    outlist = []
    arg = extract_arg(message.text)
    if len(arg) > 0:
        arg1 = arg[0].lower()
        # traverse the software list
        output = os.popen('wmic process get description, processid').read()
        # try block
        #  arrange the string
        try:
            procs = output.split('\n')
            title = procs[0]
            del procs[0]
            procs.sort()
            procs.insert(0, title)
            if arg1 != '*': outlist.append(procs[0])
            for procnm in procs:
                if procnm:
                    if arg1 in procnm.lower() or arg1 == '*':
                        outlist.append(procnm)
            send_multiline(message.chat.id, outlist)
        except IndexError as e:
            print("")
    else:
        outlist.append("/p <mask of process name>")
        outlist.append("Example: /p opera")
        send_multiline(message.chat.id, outlist, block_sz = 50)

if __name__ == '__main__':
    print("This is simpliest Remote Manager\n")
    print("/scr: Screenshot!\n")
    print("/p <mask of process name>:")
    print("   shows processes with its Ids in the system\n")
    print("/k <pid>: kill active process \n")
    print("/aw: shows ProcessId of active window (for Windows)\n")
    print("/f <pid>: set foreground process \n")
    print("/r <module name>: Run process in the system\n")
    
    bot.polling(none_stop = True, interval=0)