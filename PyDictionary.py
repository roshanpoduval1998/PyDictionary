# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import re
import json
import random
from tkinter import *
import keyboard
import tkinter as tk
import time
from tkinter import ttk
from tkinter import filedialog
from functools import partial
import random
from tkinter import *
import win32com.client

class AutoCompleteCombobox(Entry):

    """
    Inherit ENTRY from tkinter 
    and manipulating accordingly
    up & down function to move 
    in search values by the
    user
    """

    def __init__(self, word_list, *args, **kwargs):
        Entry.__init__(self, *args, **kwargs)
        self.word_list = word_list
        self.var = self["textvariable"]
        if self.var == '':
            self.var = self["textvariable"] = StringVar()
        self.var.trace('w', self.word_update)
        self.bind("<Right>", self.user_selection)
        self.bind("<Return>", self.user_selection)
        self.bind("<Up>", self.up)
        self.bind("<Down>", self.down)
        self.list_box_up = False

    def word_update(self, name, index, mode):

        #constant updation of words values when user types
        try:
            if self.var.get() == '':
                self.lb.destroy()
                self.list_box_up = False
            else:
                words = self.word_search_and_compare()
                if words:
                    if self.list_box_up == False:
                        self.lb = Listbox(width="45",bg="gray10",fg="white",font=("segoe ui",22))
                        self.lb.bind("<Double-Button-1>", self.user_selection)
                        self.lb.bind("<Right>", self.user_selection)
                        self.lb.bind("<Return>", self.user_selection)
                        self.lb.place(x=self.winfo_x(), y=self.winfo_y()+self.winfo_height())
                        self.list_box_up = True
                    self.lb.delete(0, END)
                    for w in words:
                        self.lb.insert(END,w)
                else:
                    if self.list_box_up:
                        self.lb.destroy()
                        self.list_box_up = False
        except:
            pass

    def user_selection(self, event):

        """
        for selection of search value
        by the user selection
        """
        if self.list_box_up:
            self.var.set(self.lb.get(ACTIVE))
            self.lb.destroy()
            self.list_box_up = False
            self.icursor(END)

    def up(self, event):

        if self.list_box_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            if index != '0':                
                self.lb.selection_clear(first=index)
                index = str(int(index)-1)                
                self.lb.selection_set(first=index)
                self.lb.activate(index) 

    def down(self, event):

        if self.list_box_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            if index != END:                        
                self.lb.selection_clear(first=index)
                index = str(int(index)+1)        
                self.lb.selection_set(first=index)
                self.lb.activate(index) 

    def word_search_and_compare(self):

        """
        for prediction of words
        of user typed values in entry box 
        """
        d = self.var.get()
        pattern = d
        return [w for w in word_list if re.match(pattern,w)]

# CREATE A TOOLTIP--------------------------------------------------------------------------------------------------------

class CreateToolTip(object):
    '''
    create a tooltip for a given widget
    '''
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.close)
    def enter(self, event=None):
    	time.sleep(.8)
    	x = y = 0
    	x, y, cx, cy = self.widget.bbox("insert")
    	x += self.widget.winfo_rootx() + 25
    	y += self.widget.winfo_rooty() + 30
    	# creates a toplevel window
    	self.tw = tk.Toplevel(self.widget)
    	# Leaves only the label and removes the app window
    	self.tw.wm_overrideredirect(True)
    	self.tw.wm_geometry("+%d+%d" % (x, y))
    	self.tw.wm_attributes('-alpha',.80)
    	label = tk.Label(self.tw, text=self.text, justify='left',
                       background='black', fg="white", relief='solid', borderwidth=1,
                       font=("segoe ui", "14", "normal"))
    	label.pack(ipadx=1)
    def close(self, event=None):
        if self.tw:
            self.tw.destroy()

# OBJECT : SEARCH_CLASS---------------------------------------------------------------------------------------------------

class dict_search__:

    def __init__(self,Document):
        windowWidth = 1095
        windowHeight = 700

        root.geometry("1095x700")
        x = entry.get().strip().lower()
        self.x = x
        y = data[x]
        self.y = y
        r = x.capitalize()
        self.r = r
        remove_extra_space_1 = y.split(". ")
        merge_new_list_1 = ".\n".join(remove_extra_space_1)
        remove_extra_space_2 = merge_new_list_1.split(".\n.")
        merge_new_list_2 = "\n".join(remove_extra_space_2)
        ordered_paragraph = ("\n\n{}-\n\n{}".format(r,merge_new_list_2))
        self.word_meaning = ordered_paragraph
        wordtovoice = open("{}\\cachevoice__.txt".format(Document),"w")
        wordtovoice.write("{},".format(self.r))
        wordtovoice.close()
        wordscalled = open("{}\\cachecall__.txt".format(Document),"+a")
        wordscalled.write("{},".format(self.r))
        wordscalled.close()

    def FindWord(self):
        
        try:
            text = tk.Text(root,)
            text.insert(END, self.word_meaning)
            text.config(width="95", bg="black", fg="white",font=("segoe ui",13,"italic")
                ,borderwidth=0,highlightthickness=0,selectbackground="gray10")
            text.place(x=10, y=100)
            text.config(state=DISABLED)
        except KeyError:
            messagebox_show("Error","Enter a valid word")

    def save_word(self):

        try:
            output_file = filedialog.asksaveasfile(mode="w", defaultextension='txt'
                , filetypes=(("Text File","*txt*"),("Microsoft Word File","*docx*"),("All files","*.*")))
            word_save = open("{}".format(output_file.name,self.r),"w")
            word_save.write(self.word_meaning)
            word_save.close()
        except:
            pass

    def printer(self):
        try:
        	output_file = filedialog.asksaveasfile(mode="w", defaultextension='txt'
	            , filetypes=(("Text File","*txt*"),("Microsoft Word File","*docx*"),("All files","*.*")))
        	word_save = open("{}".format(output_file.name),"w")
        	word_save.write(self.word_meaning)
        	word_save.close()
        	os.startfile("{}".format(output_file.name), "print")
        except:
        	pass

# OBJECT : RANDOM_WORDS -----------------------------------------------------------------------------------------------------

class random__:

    def __init__(self,Document):
        windowWidth = 1095
        windowHeight = 700

        root.geometry("1095x700")
        v = random.choice(word_list)
        self.v = v
        r = v.capitalize()
        wordtovoice = open("{}\\cachevoice__.txt".format(Document),"w")
        wordtovoice.write("{},".format(r))
        wordtovoice.close()
        wordscalled = open("{}\\cachecall__.txt".format(Document),"+a")
        wordscalled.write("{},".format(r))
        wordscalled.close()
        word_shuffle = data[v]
        remove_extra_space_1 = word_shuffle.split(". ")
        merge_new_list_1 = ".\n".join(remove_extra_space_1)
        remove_extra_space_2 = merge_new_list_1.split(".\n.")
        merge_new_list_2 = "\n".join(remove_extra_space_2)
        ordered_paragraph = ("\n\n{}-\n\n{}".format(r,merge_new_list_2))
        self.ordered_paragraph = ordered_paragraph

    def random_words(self):
        
        try:
            entry.insert(END, self.v)
            text = tk.Text(root,)
            text.insert(END, self.ordered_paragraph)
            text.config(width="95", bg="black", fg="white",font=("segoe ui",13,"italic")
                ,borderwidth=0,highlightthickness=0,selectbackground="gray10")
            text.place(x=10, y=99)
            text.config(state=DISABLED)
        except KeyError:
            messagebox_show("Error","Enter a valid word")

# OBJECT : HISTORY ----------------------------------------------------------------------------------------------------------

class Histroy__:

    def __init__(self,Document):

        root.geometry("1280x700")
        word_history = open("{}\\cachecall__.txt".format(Document),"r")
        list_x = word_history.readline(-1)
        word_history.close()
        list_word = list_x.split(",")
        self.list_word = list_word[::-1]

    def app_history(self):
        
        history = tk.Label(root,text="Recently searched", bg="black", fg="white"
            ,font=("segoe ui",10,"italic"),borderwidth=0,highlightthickness=0)
        history.place(x=1100,y=50)    
        text = tk.Text(root,)
        text.config(width="90", height="30", bg="black", fg="white",borderwidth=0,
            highlightthickness=0,selectbackground="gray1")
        text.place(x=1100, y=92)
 
        for i in self.list_word:
            def func(i):
                entry.delete(0,END)
                entry.insert(END,i)
            button = Button(root, text=i, bg="black", fg="white",highlightthickness=0
                ,font=("segoe ui",10,"italic"),borderwidth=0,command=partial(func, i))
            text.window_create("end", window=button)
            text.insert("end", "\n")

class messagebox_show:

	def __init__(self,message1,message2):
		try:
			error_box = tk.Toplevel()
			error_box.bell()
			windowWidth = 200
			windowHeight = 100
			x_ = int(error_box.winfo_screenwidth()/2 - windowWidth/2)
			y_ = int(error_box.winfo_screenheight()/2 - windowHeight/2)
			error_box.geometry("200x100+{}+{}".format(x_,y_))
			error_box.iconbitmap("{}/icons/main_icon.ico".format(loco))
			error_box.configure(bg="gray20")
			error_box.title(message1)
			label = tk.Label(error_box, text=message2,bg="gray20",fg='white',font=("Banschrift",13))
			label.place(x=30,y=20)
			button = tk.Button(error_box, text="OK",width="15", bg="gray80",fg='gray1',borderwidth=0,highlightthickness=0,command=self.destroy_button)
			button.place(x=50,y=70)
			root.wm_attributes('-alpha',0.7)
			entry.config(state=DISABLED)
			error_box.wm_attributes('-alpha',0.95)
			error_box.after(1, lambda: error_box.focus_force())
			error_box.bind("<Return>",self.destroy_button_Event)
			if root == False:
				error_box.quit()
			error_box.protocol('WM_DELETE_WINDOW',self.destroy_button)
			error_box.resizable(width=0,height=0)
			self.error_box = error_box
			self.root = root
			self.entry = entry
			error_box.mainloop()
		except:
			pass

	def exit__(self):
		self.error_box.destroy()

	def destroy_button(self):
		self.error_box.destroy()
		try:
		    self.root.wm_attributes('-alpha',0.92)
		    self.entry.config(state=NORMAL)
		except:
		    pass
	def destroy_button_Event(self,Event):
		self.error_box.destroy()
		try:
			self.root.wm_attributes('-alpha',0.92)
			self.entry.config(state=NORMAL)
		except:
			pass

class instructionbox_show:

	def __init__(self,message1,message2,message3):
		try:
			instruction_box = tk.Toplevel()
			instruction_box.bell()
			windowWidth = 200
			windowHeight = 100
			x_ = int(instruction_box.winfo_screenwidth()/2 - windowWidth/2)
			y_ = int(instruction_box.winfo_screenheight()/2 - windowHeight/2)
			instruction_box.geometry("+{}+{}".format(x_,y_))
			instruction_box.iconbitmap("{}/icons/main_icon.ico".format(loco))
			instruction_box.configure(bg="gray20")
			instruction_box.title(message1)
			label = tk.Label(instruction_box, text=message2,bg="gray20",fg='white',font=("segoe ui",13))
			label.grid(row=0)
			button = tk.Button(instruction_box, text=message3, font=("segoe ui",13), bg="gray70",fg='gray1', command=self.github_,borderwidth=0,highlightthickness=0)
			button.grid(row=1)
			label = tk.Label(instruction_box, text="\n",bg="gray20",fg='white')
			label.grid(row=2)
			button = tk.Button(instruction_box, text="OK",width="15", bg="gray80",fg='gray1',borderwidth=0,highlightthickness=0,command=self.destroy_button)
			button.grid(row=3)
			label = tk.Label(instruction_box, text="\n",bg="gray20",fg='white')
			label.grid(row=4)
			instruction_box.after(1, lambda: instruction_box.focus_force())
			instruction_box.bind("<Return>", self.destroy_button_Event)
			instruction_box.protocol('WM_DELETE_WINDOW',self.destroy_button)
			instruction_box.resizable(width=0,height=0)
			self.instruction_box = instruction_box
			self.root = root
			self.entry = entry
			instruction_box.mainloop()
		except:
			pass

	def github_(self):
		os.startfile("https://www.github.com/roshanpoduval1998/PyDictionary")

	def exit__(self):
		self.instruction_box.destroy()

	def destroy_button(self):
		self.instruction_box.destroy()

	def destroy_button_Event(self,Event):
		self.instruction_box.destroy()

class Text_to_Speak__:

    def __init__(self,Document):
        
        wordtovoice = open("{}\\cachevoice__.txt".format(Document),"r")
        x = wordtovoice.readline(-1)
        wordtovoice.close()
        self.x = x

    def get_word(self):

        y = self.x.split(",")
        speak = y[0]
        speak_word = speak[0:]
        return str(speak_word)

    def speak_word(self):
        
        if str(self.x) == '':
            speaker = win32com.client.Dispatch("SAPI.SpVoice") 
            s = str(self.get_word())
            speaker.Speak(s)
        else:
            speaker = win32com.client.Dispatch("SAPI.SpVoice") 
            s = self.x
            speaker.Speak(s)

def copy_content():
	pass

def select_all_content():
	pass

"""CREATING 2 .txt
FILES FOR VOICE
AND WORD CALLED"""

def Extractcache__(Document):

    # CACHE-CREATION-IN-TEXT-FORM----------------------

    wordtovoice = open("{}\\cachevoice__.txt".format(Document),"w")
    wordtovoice.write("")
    wordtovoice.close()

    words_called = open("{}\\cachecall__.txt".format(Document),"w")
    words_called.write("")
    words_called.close()

# DRIVER-FUNCTIONS------------------------------------------------------------------------------

# TO FIND MEANING IN DICTIONARY USING OBJECT dict_search__

def search_class():
    
    try:
    	if entry.get() != "":
    		a = dict_search__(Document)
    		return a.FindWord()
    except:
    	messagebox_show("Error","Enter a valid word")

def search_class_event(Event):
    
    try:
    	if entry.get() != "":
    		a = dict_search__(Document)
    		return a.FindWord()
    except:
    	messagebox_show("Error","Enter a valid word")

# SAVING THE WORDS MEANING

def save_class():

    try:
        a = dict_search__(Document)
        return a.save_word()
    except:
        messagebox_show("Error","Enter a valid word")

def save_class_event(Event):

    try:
        a = dict_search__(Document)
        return a.save_word()
    except:
        messagebox_show("Error","Enter a valid word")

# OPENS PRINTER

def printer():
    
    try:
        a = dict_search__(Document)
        return a.printer()
    except:
        messagebox_show("Error","Unable to call printer")

def printer_event(Event):
    
    try:
        a = dict_search__(Document)
        return a.printer()
    except:
        messagebox_show("Error","Unable to call printer")

# DISPLAY RANDOM WORDS MEANING USES OBJECT random__

def randomness():
    
    a = random__(Document)
    entry.delete(0,END)
    return a.random_words()

def randomness_event(Event):
    
    a = random__(Document)
    entry.delete(0,END)
    return a.random_words()

# OPENS HISTORY TAB USES OBJECT HISTORY

def open_history():
    
    a = Histroy__(Document)
    return a.app_history()

def open_history_event(Event):
    
    a = Histroy__(Document)
    return a.app_history()


# TEXT TO SPEECH FOR WORDS

def speak_word_driver():

    if entry.get() != '':
        a = Text_to_Speak__(Document)
        return a.speak_word()
    else:
        pass

def speak_word_driver_event(Event):

    if entry.get() != '':
        a = Text_to_Speak__(Document)
        return a.speak_word()
    else:
        pass

# CLEARS ANY VALUE IN ENTRY

def clear_entry():

    windowWidth = 1095
    windowHeight = 700

    root.geometry("1095x700")
    entry.delete(0,END)
    text = tk.Text(root,)
    text.insert(END, "")
    text.config(width="95", bg="black", fg="white",font=("segoe ui",13,"italic")
        ,borderwidth=0,highlightthickness=0,selectbackground="gray10")
    text.place(x=10, y=100)
    text.config(state=DISABLED)
    entry.focus()

def clear_entry_event(Event):

    windowWidth = 1095
    windowHeight = 700

    root.geometry("1095x700")
    entry.delete(0,END)
    text = tk.Text(root,)
    text.insert(END, "")
    text.config(width="95", bg="black", fg="white",font=("segoe ui",13,"italic")
        ,borderwidth=0,highlightthickness=0,selectbackground="gray10")
    text.place(x=10, y=100)
    text.config(state=DISABLED)
    entry.focus()

# DISPLAYS SUPPORT FOR THE APP

def info_help():
	instruction = "Save - Ctrl+S\nPrint - Ctrl+P\nInstructions - Ctrl+I\nClear - Ctrl+C or Delete\nSpeak - Ctrl+Q\nHistory - Ctrl+H\nRandom Word - Ctrl+R\nSubmit Word - Enter\n"
	instructionbox_show("i",instruction,"Info?")
	
def info_help_event(Event):
	instruction = "Save - Ctrl+S\nPrint - Ctrl+P\nInstructions - Ctrl+I\nClear - Ctrl+C or Delete\nSpeak - Ctrl+Q\nHistory - Ctrl+H\nRandom Word - Ctrl+R\nSubmit Word - Enter\n"
	instructionbox_show("i",instruction,"Info?")


if __name__ == "__main__":

    # INITIALIZE-GUI----------------------------------------------------------------------------------

    root = tk.Tk()
    windowWidth = 1095
    windowHeight = 700
    x_ = int(root.winfo_screenwidth()/2 - windowWidth/2)
    y_ = int(root.winfo_screenheight()/2 - windowHeight/2)
    user_name = os.getlogin()
    Document = "C:\\Users\\{}\\Documents".format(user_name)
    try:
    	os.mkdir(Document+"\\PyDictionary")
    except:
    	pass
    Document = Document+"\\PyDictionary"
    loco = os.path.dirname(os.path.realpath(__file__))
    Extractcache__(Document)
    root.geometry("1095x700+{}+{}".format(x_,y_))
    root.configure(bg="black")
    root.title("Py Dictionary")
    root.bind("<Return>",search_class_event)
    root.bind("<Control-s>", save_class_event)
    root.bind("<Control-p>", printer_event)
    root.bind("<Control-r>", randomness_event)
    root.bind("<Control-h>", open_history_event)
    root.bind("<Control-q>", speak_word_driver_event)
    root.bind("<Control-c>", clear_entry_event)
    root.bind("<Delete>", clear_entry_event)
    root.bind("<Control-i>", info_help_event)

    # INITIALIZE-JSON-FILE-----------------------------------------------------------------------------------------------------------------------

    with open('{}/JSON/DATA__.json'.format(loco),'r') as json_file:
    	data = json.load(json_file)
    word_list = [*data]

    # OPERATORS----------------------------------------------------------------------------------------------------------------------------------

    try:
        entry = AutoCompleteCombobox(word_list, root,)
        entry.config(width="52",bg="gray10",fg="white"
            ,font=("segoe ui",25),insertbackground="white")
        entry.focus_set()
        entry.place(x=80,y=50)
    except:
        entry = tk.Entry(width="52",bg="gray10",fg="white"
            ,font=("segoe ui",25),insertbackground="white")
        entry.focus_set()
        entry.place(x=80,y=50)
    root.bind("<Double-Button-1>",search_class_event)
    #-------------------------------------------------------------------------------------------------

    text = tk.Text(root,)
    text.config(width="95", bg="black", fg="white",font=("segoe ui",13)
        ,borderwidth=0,highlightthickness=0,selectbackground="gray10")
    text.place(x=10, y=100)
    text.config(state=DISABLED)
    #-------------------------------------------------------------------------------------------------

    clear = tk.Button(root, text="X", width="2",bg="gray10",fg="white"
        , font=("segoe ui",15)
        ,borderwidth=0,highlightthickness=0,)
    clear.place(x=980,y=53)
    clear.bind("<Button-1>",clear_entry_event)

    #-------------------------------------------------------------------------------------------------
    try:
    	infoimg = tk.PhotoImage(file="{}/buttons/instruction.png".format(loco))
    	info = tk.Button(root, image=infoimg
    		, command=info_help, highlightthickness=0, borderwidth=0)
    	info.place(x=185,y=5)
    except:
    	info = tk.Button(root, text=' i', width="2",bg="gray1",fg="white"
    		, command=info_help, font=("segoe ui",14,"bold")
    		, highlightthickness=0, borderwidth=0,anchor=W)
    	info.place(x=185,y=5)

    #---------------------------------------------------------------------------------------------------

    try:
        submitimg = tk.PhotoImage(file="{}/buttons/submit.png".format(loco))
        submit = tk.Button(root, image=submitimg,highlightthickness=0,borderwidth=0
            ,command=search_class)
        submit.place(x=1030,y=52)
    except:
        submit = tk.Button(root, text="→", width="3",bg="gray10",fg="white"
            , command=search_class, font=("segoe ui",16))
        submit.place(x=1030,y=52)

    #-------------------------------------------------------------------------------------------------
        
    try:
        saveimg = tk.PhotoImage(file="{}/buttons/save.png".format(loco))
        save_button = tk.Button(root, image=saveimg,borderwidth=0,highlightthickness=0
            ,command=save_class)
        save_button.place(x=5,y=5)
    except:
        save_button = tk.Button(root, text="↓", bg="black", fg="white", command=save_class
            , font=("segoe ui",15,"bold"), borderwidth=0, highlightthickness=0)
        save_button.place(x=5,y=5)

    #-------------------------------------------------------------------------------------------------
        
    try:
        speakimg = tk.PhotoImage(file="{}/buttons/speak.png".format(loco))
        speak_word = tk.Button(root, image=speakimg,highlightthickness=0,borderwidth=0
            , command=speak_word_driver)
        speak_word.place(x=30,y=57)
    except:
        speak_word = tk.Button(root, text='♪', width="3",bg="gray1",fg="white"
            , command=speak_word_driver, font=("segoe ui",16))
        speak_word.place(x=30,y=57)
    #-------------------------------------------------------------------------------------------------
        
    try:
        randomimg = tk.PhotoImage(file="{}/buttons/random.png".format(loco))
        random_word = tk.Button(root, image = randomimg,highlightthickness=0
            ,borderwidth=0, command=randomness)
        random_word.place(x=95,y=5)
    except:
        random_word = tk.Button(root, text='⌂', width="3",bg="gray1",fg="white"
            , command=randomness, font=("segoe ui",16), highlightthickness=0, borderwidth=0,)
        random_word.place(x=95,y=5)
    #-------------------------------------------------------------------------------------------------
        
    try:
        historyimg = tk.PhotoImage(file="{}/buttons/history.png".format(loco))
        history = tk.Button(root, image = historyimg,highlightthickness=0
            ,borderwidth=0, command=open_history)
        history.place(x=140,y=5)
    except:
        history = tk.Button(root, text='⌂', width="3",bg="gray1",fg="white"
            , command=open_history, font=("segoe ui",16), highlightthickness=0, borderwidth=0,)
        history.place(x=140,y=5)
    #-------------------------------------------------------------------------------------------------
        
    try:
        printerimg = tk.PhotoImage(file="{}/buttons/printer.png".format(loco))
        printer = tk.Button(root, image = printerimg,highlightthickness=0
            ,borderwidth=0, command=printer)
        printer.place(x=50,y=5)
    except:
        printer = tk.Button(root, text='⌂', width="3",bg="gray1",fg="white"
            , command=printer, font=("segoe ui",16), highlightthickness=0, borderwidth=0,)
        printer.place(x=50,y=5)

	# Create Tool-Tip---------------------------------------------------------------------------------------

    CreateToolTip(printer, "Print\nCtrl+P")
    CreateToolTip(save_button, "Save\nCtrl+S")
    CreateToolTip(random_word, "Random Word\nCtrl+R")
    CreateToolTip(history, "Histroy\nCtrl+H")
    CreateToolTip(speak_word, "Speak\nCtrl+Q")
    CreateToolTip(clear, "Clear\nCtrl+C")
    CreateToolTip(info, "Instructions\nCtrl+I")

    #---------------------------------------------------------------------------------------------------

    entry.focus_force()
    root.resizable(width=0,height=0)
    root.wm_attributes('-alpha',0.92)
    root.protocol("WM_DELETE_WINDOW",root.destroy)
    root.iconbitmap("{}/icons/main_icon.ico".format(loco))
    root.mainloop()