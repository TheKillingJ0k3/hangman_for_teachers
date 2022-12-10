#! python3

from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import openpyxl
from openpyxl.styles import Font, Border, Side
import os, shutil


#GLOBAL VARIABLES
# var = ''
# last_name_initial_var = ''
level_selected = ''
number_of_letters = ''
attempts = ''

#####################################################
# TODO: gaming section: player chooses level, words of this level are saved in a list, program chooses randomly a word from this list, game begins
# TODO: exclude articles from the above
# TODO: languages?

##################################  FUNCTIONS  ##################################################

#creates  folder inside wd, if it doesn't already exist
def createFolder(path):
    try:
        if not os.path.exists(path):
            os.mkdir(path)
    except OSError:
        print('Error creating directory' + path)

################## styles ##################
def excel_styles(excel_sheet):
    Arial_11_Font = Font(name='Arial', size=11)

    Arial_11_bold_Font = Font(name='Arial', size=11, bold=True)
    for columnNum in range(1, excel_sheet.max_column + 1):
        excel_sheet.cell(row=1, column=columnNum).font = Arial_11_bold_Font
    excel_sheet.freeze_panes = 'A2'
############################################

def open_excel(excel_file_path):
    try:
        wb = openpyxl.load_workbook(excel_file_path, data_only=True) # , data_only=True in case file has a lot of formulas
        ws = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'A1' # or cell.value = skata
        ws.column_dimensions['A'].width = 22.43
        ws['B1'] =  'A2'
        ws.column_dimensions['B'].width = 22.43
        ws['C1'] =  'B1'
        ws.column_dimensions['C'].width = 22.43
        ws['D1'] =  'B2'
        ws.column_dimensions['D'].width = 22.43
        ws['E1'] =  'C1'
        ws.column_dimensions['E'].width = 22.43
        ws['F1'] =  'C2'
        ws.column_dimensions['F'].width = 22.43
        excel_styles(ws)
    wb.save(excel_file_path)
    return wb, ws

######################################################
def show_words():
    os.startfile(r".\\Hangman Excel\\Hangman Excel.xlsx") # , data_only=True in case file has a lot of formulas

def select_level(): #connect selection in GUI with a variable
    pass

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
createFolder('.\\Hangman Excel')
wb, ws = open_excel('.\\Hangman Excel\Hangman Excel.xlsx')



###################### main GUI - Button creation #########################################
root = Tk()
root.title('Hangman')
root.geometry('500x350')
# root.state('zoomed')
# root.option_add('*tear0ff', False) #opens fullscreen
# root.iconbitmap('.\\hangman.ico')

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
menubar.add_cascade(label='New Game', menu=action_menu) #creates name of new submenu
level_menu = Menu(action_menu, tearoff=False)
level_menu.add_command(label='A1', command=function) #adds option to submenu
level_menu.add_command(label='A2', command=function) #adds option to submenu
level_menu.add_command(label='B1', command=function) #adds option to submenu
level_menu.add_command(label='B2', command=function) #adds option to submenu
level_menu.add_command(label='C1', command=function) #adds option to submenu
level_menu.add_command(label='C2', command=function) #adds option to submenu
action_menu.add_cascade(label='Choose Level', menu=level_menu) #creates name of new submenu

settings_menu = Menu(menubar, tearoff=False)
menubar.add_cascade(label='Settings', menu=settings_menu) #creates name of new submenu
settings_menu.add_command(label='Show Saved Words', command=show_words) #  main update_itineraries
word_settings_menu = Menu(settings_menu, tearoff=False)
settings_menu.add_cascade(label='Add Words to a Level', menu=word_settings_menu) #creates name of new submenu
word_settings_menu.add_command(label='A1', command=function) #adds option to submenu
word_settings_menu.add_command(label='A2', command=function) #adds option to submenu
word_settings_menu.add_command(label='B1', command=function) #adds option to submenu
word_settings_menu.add_command(label='B2', command=function) #adds option to submenu
word_settings_menu.add_command(label='C1', command=function) #adds option to submenu
word_settings_menu.add_command(label='C2', command=function) #adds option to submenu


advice_menu = Menu(menubar, tearoff=False) #first_lineises new submenu
menubar.add_cascade(label='Advice', menu=advice_menu) #creates name of new submenu
advice_menu.add_command(label='1', command=function) #adds option to submenu
# menu_app.add_command(Label='Save app', command=save_excel) #adds option to submenu

menu_exit = Menu(menubar)
menubar.add_cascade(label='Exit', command=frame.quit)
######################################################################################

root.mainloop()