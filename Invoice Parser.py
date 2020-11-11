from tkinter import *
from tkinter import ttk 
from json import loads, dumps
import configparser
import xmltodict
import xlsxwriter
import sys
import os
import timeit
from progress.bar import Bar
import logging
from Invoice import Invoice

def get_files(path):
    f = []
    for (dirpath, dirnames, filenames) in os.walk(path):
        f.extend(filenames)
        break
    return f

def load_config():
    config = configparser.ConfigParser()
    config.read("config.ini")
    return config
    
def run():
    starttime = timeit.default_timer()
    config = load_config()
    files = get_files(config["files"]["source_path"])
    fails = 0
    if files:
        for file in files:
            test = Invoice(file, config)
            try:
                new_file = test.create_file()
                logger.info("Datei: "+ file +" wurde bearbeitet. Datei "+ new_file +" wurde erstellt")
            except:
                logger.debug("Fehler beim Bearbeiten der Datei: "+file)
            try:
                test.cleanup()
            except Exception:
                fails += 1
        erg = "Dateien Verarbeitet:" + str(len(files))+ "\nDavon fehlerhaft:"+ str(fails)+ "\nBenötigte Zeit: %.2f Sekunden" % (timeit.default_timer() - starttime) 
        logger.info(erg.replace('\n',' |'))
    else:
        erg = "Keine Dateien im Verzeichnis \n" + config["files"]["source_path"] + "\ngefunden."
        logger.debug(erg.replace('\n',''))
    # print output to "Ergebnis"-textfield
    e4.config(state=NORMAL)
    e4.delete(1.0, END)
    e4.insert(10.0, erg)
    e4.config(state=DISABLED)

    
def save():
    # This saves only the config file. NOT the .xml files!
    config["files"] = {}
    # puts "/" at the end of the path if it isn't there
    config["files"]["source_path"] = e1.get("1.0","end-1c")
    if config["files"]["source_path"][-1] != "\\":
        config["files"]["source_path"] += "\\"
    config["files"]["destination_path"] = e2.get("1.0","end-1c")
    if config["files"]["destination_path"][-1] != "\\":
        config["files"]["destination_path"] += "\\"
    config["usk"] = {}
    config["usk"]["USK_Liste"] = e3.get("1.0","end-1c")
    try:
        with open("config.ini", "w") as configfile:
            config.write(configfile)
            logger.info("Config Datei wurde geändert: Source_path: "+ config["files"]["source_path"]+" | Destination_path: "+ config["files"]["destination_path"] +" | USK_Liste:"+ config["usk"]["USK_Liste"].replace('\n',''))
    except Esception as e:
        logging.error(traceback.format_exc())
        logger.debug("Config Datei konnte nicht gespeichert werden")
    refresh()


def refresh():
    config.read("config.ini")
    try:
        source = config["files"]["source_path"]
    except:
        source = ""
    try:
        dest = config["files"]["destination_path"]
    except:
        dest = ""
    try:
        conf = config["usk"]["USK_Liste"]
    except:
        conf = ""
    try:
        temp = erg
    except:
        temp = ""
    e1.delete(1.0,END)
    e2.delete(1.0,END)
    e3.delete(1.0,END)
    #e4.delete(1.0,END)
    e1.insert(10.0, source)
    e2.insert(10.0, dest)
    e3.insert(10.0, conf)
    # print(temp)
    # e4.insert(10.0, temp)
    usk_lock()


def usk_list(data, entry_row, row, col):
    row_temp = row
    col_temp = col
    for key, val in data.items():
        # this should only occure if .olav USK is processed
        if type(val) is dict:
            col_temp +=1
            usk_list(val,entry_row,row_temp,col_temp)
        else:
            entry_row[key] = Entry(master)
            entry_row[val] = Entry(master)
            entry_row[key].insert(10,key)
            entry_row[val].insert(10,val)
            entry_row[key].grid(row=3+row_temp, column=1+col_temp)
            entry_row[val].grid(row=3+row_temp, column=2+col_temp)
            row_temp +=1
    col_temp -=1
    row_temp +=1
    return entry_row

def usk_lock():
    # disables the USK textfield
    if lock.get() == True:
        e1.config(state=DISABLED)
        e2.config(state=DISABLED)
        e3.config(state=DISABLED)
    elif lock.get() == False :
        e1.config(state=NORMAL)
        e2.config(state=NORMAL)
        e3.config(state=NORMAL)


# create logger
logging.basicConfig(handlers=[logging.FileHandler(filename="./log.txt", 
                                                 encoding='utf-8', mode='a+')],
                    format="%(asctime)s - %(name)s - %(funcName)s - %(lineno)s - %(levelname)s - %(message)s",
                    level=logging.INFO)
logger = logging.getLogger("Invoice Parser")


config = configparser.ConfigParser()
config.read("config.ini")
output = ""
erg = ""
window = Tk()
window.title("Invoice Parser - XML to XLSX")
# window.geometry("800x530")
lock = BooleanVar()
lock.set(True)

# Creates the images on the UI
# temporarily removed because ugly
# filename = PhotoImage(file="bg.gif")
# background_label = ttk.Label(window, image=filename)
# background_label.place(x=0, y=0, relwidth=1, relheight=1)

# create UI level
master = ttk.Frame(window)
# master.config(bg="white")
master.grid(row=0,column=0)

# Creates all lables on the UI
ttk.Label(master, text="source_path", borderwidth=10).grid(row=0)
ttk.Label(master, text="output_path", borderwidth=10).grid(row=1)
ttk.Label(master, text="USK", borderwidth=10).grid(row=2)
ttk.Label(master, text="Ergebnis", borderwidth=10).grid(row=3)

# Textfields on the UI
e1 = Text(master, height=1, width=40)
e2 = Text(master, height=1, width=40)
e3 = Text(master, height=10, width=40)
e4 = Text(master, height= 4, width=40)

refresh()

# Aligns the Textobjects
e1.grid(row=0, column=1, columnspan=2)
e2.grid(row=1, column=1, columnspan=2)
e3.grid(row=2, column=1, columnspan=2)
e4.grid(row=3, column=1, columnspan=2)
e4.config(state=DISABLED)

# Buttons
ttk.Button(master, text='Beenden', command=master.quit).grid(row=99, column=0, sticky=W, pady=4)
ttk.Button(master, text='Speichern', command=save).grid(row=99, column=1, sticky=W, pady=4)
ttk.Button(master, text='Ausführen', command=run).grid(row=99, column=3, sticky=W, pady=4)

# Checkbox for activating/deactivating the USK textfield

c = Checkbutton(master, text="Lock Settings", variable=lock, command=usk_lock).grid(row=99, column=2, sticky=W, pady=4)

# sets the veriables for the mainfunction
row=0
col=0
try:
    usk = loads(config["usk"]["usk_liste"])
except:
    usk = {}
entry_row = {}

window.mainloop( )
