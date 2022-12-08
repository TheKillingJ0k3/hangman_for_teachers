#! python3

from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from openpyxl import *
import os, shutil


#GLOBAL VARIABLES
var = ''
last_name_initial_var = ''

#####################################################


##################################  FUNCTIONS  ##################################################

#creates  folder inside wd, if it doesn't already exist
def createFolder(path):
    try:
        if not os.path.exists(path):
            os.mkdir(path)
    except OSError:
        print('Error creating directory' + path)

######################################################

def function():
    pass

def function1():
    pass

# def save_variable(): #saves var when input button is pressed
#     global last_name_initial_var
#     last_name_initial_var = last_name_entry_initial.get()
#     print (last_name_initial_var)
# # as an example, this function is used in last_name_initial_submit
# # when user writes last name and presses submit, script prints var's value

#############################################################################################

###################### main GUI - Button creation #########################################
root = Tk()
root.title('Hangman')
root.geometry('500x350')
# root.state('zoomed')
# root.option_add('*tear0ff', False) #opens fullscreen

# background_image = PhotoImage(file='C:\\Users\\kj\\Documents\\Python Projects\\Comic downloader\\crowd-img.png')
# background_label = Label(root, image=background_image)
# background_label.place(x=0, y=0, relwidth=1, relheight=1)

frame = Frame(root, borderwidth=5, relief="sunken", width=500, height=500) # 100 -200
frame.pack()

##############################   MENU  #############################################
menubar = Menu(root) #creates menubar
root.config(menu = menubar) #same as frame['menu'] = menubar, doesn't need menu=menu_file etc inside cascade

#creating submenus in frame menu/menubar
data_menu = Menu(menubar, tearoff=False) #first_lineises new submenu
# menubar.add_cascade(label='View Data', menu=data_menu) #creates name of new submenu
# data_menu.add_command(label='Buses', command=open_Buses_window) # if command function has argument, window opens automatically


action_menu = Menu(menubar, tearoff=False) #first_lineises new submenu
menubar.add_cascade(label='Actions', menu=action_menu) #creates name of new submenu
metrics_menu = Menu(action_menu, tearoff=False)
metrics_menu.add_command(label='1', command=function) #adds option to submenu
action_menu.add_cascade(label='1', menu=metrics_menu) #creates name of new submenu

action_menu.add_command(label='1', command=function) #  main update_itineraries
action_menu.add_command(label='2', command=function) #  main update_itineraries

plto_menu = Menu(action_menu, tearoff=False)
# plto_menu.add_command(label='Update PLTOs', command=PLTOs) #adds option to submenu
plto_menu.add_command(label='1', command=function) #adds option to submenu
action_menu.add_cascade(label='1', menu=plto_menu) #creates name of new submenu

advice_menu = Menu(menubar, tearoff=False) #first_lineises new submenu
menubar.add_cascade(label='Advice', menu=advice_menu) #creates name of new submenu
advice_menu.add_command(label='1', command=function) #adds option to submenu
# menu_app.add_command(Label='Save app', command=save_excel) #adds option to submenu

menu_exit = Menu(menubar)
menubar.add_cascade(label='Exit', command=frame.quit)
######################################################################################

root.mainloop()