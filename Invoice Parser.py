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


class Invoice:

    def __init__(self, file, config):
        self.filepath = config["files"]["source_path"]+file
        with open(self.filepath, "r", encoding="utf-8") as fd:
            self.doc = xmltodict.parse(fd.read())
            self.FileSender = self.doc["epay21Finance"]["PSPData"]["FileSender"]
            self.FileName = self.doc["epay21Finance"]["PSPData"]["FileName"]
            # self.OutputFile = config["files"]["destination_path"]+self.FileName.split(".")[0]+".xlsx"
            self.OutputFile = config["files"]["destination_path"]+file.split(".")[0]+".xlsx"
            self.FileTimestamp = self.doc["epay21Finance"]["PSPData"]["FileTimestamp"]
            self.PeriodFrom = self.doc["epay21Finance"]["PSPData"]["PeriodFrom"]
            self.PeriodTo = self.doc["epay21Finance"]["PSPData"]["PeriodTo"]
            self.Amount = self.doc["epay21Finance"]["PSPData"]["Amount"]
            self.Currency = self.doc["epay21Finance"]["PSPData"]["Currency"]
            if "Purpose" in self.doc["epay21Finance"]["PSPData"]:
                self.Purpose = self.doc["epay21Finance"]["PSPData"]["Purpose"]
            else:
                self.Purpose = "n/a"
            self.columns = ["Verfahren", "USK", "Betrag", "Währung", "Einzahler", "Verwendungszweck", "Zeitstempel", "Bezahlmethode"]
            self.RecordEntry = self.doc["epay21Finance"]["Records"]["RecordEntry"]

    def pprint(self):
        print(dumps(self.doc, indent=4, sort_keys=True))

    def get_USK(self, RecordEntry):
        usk_config = loads(config["usk"]["usk_liste"].replace("\'","\""))
        usk = usk_config[RecordEntry["epay21App"]]
        if RecordEntry["epay21App"] == "hsh.olav":
            usk = usk_config["hsh.olav"][RecordEntry["Purpose"].split("/")[0]]
        return usk

    def create_table(self,sheet, row, RecordEntry):
        sheet.write(row, 0, RecordEntry["epay21App"])
        sheet.write(row, 1, self.get_USK(RecordEntry))
        sheet.write(row, 2, RecordEntry["Amount"])
        sheet.write(row, 3, RecordEntry["Currency"])
        sheet.write(row, 4, RecordEntry["PayerInfo"])
        sheet.write(row, 5, RecordEntry["Purpose"])
        sheet.write(row, 6, RecordEntry["Timestamp"])
        sheet.write(row, 7, RecordEntry["PayMethod"])

    def create_file(self):
        try:
            workbook = xlsxwriter.Workbook(self.OutputFile)
        except:
            logger.debug(self.OutputFile+(" konnte nicht erstelt werden"))
        
        sheet1 = workbook.add_worksheet("Informationen")
        sheet2 = workbook.add_worksheet("Daten")

        basicdata = (
            ["FileSender", self.FileSender],
            ["FileName", self.FileName],
            ["FileTimestamp", self.FileTimestamp],
            ["PeriodFrom", self.PeriodFrom],
            ["PeriodTo", self.PeriodTo],
            ["Amount", self.Amount],
            ["Currency", self.Currency],
            ["Purpose", self.Purpose],
        )

        row = 0
        for key, value in (basicdata):
            sheet1.write(row, 0,     key)
            sheet1.write(row, 1, value)
            row += 1

        sheet1.set_column(0, 0, 15)
        sheet1.set_column(1, 1, 70)

        for name in self.columns:
            sheet2.write(0, self.columns.index(name), name)

        row = 1
        if type(self.RecordEntry) is list:
            for entry in self.RecordEntry:
                self.create_table(sheet2, row, entry)
                row += 1
        else:
            self.create_table(sheet2, row, self.RecordEntry)

        sheet1.set_column(0, 0, 15)
        sheet1.set_column(1, 1, 70)
        sheet2.set_column(0, 0, 15)
        sheet2.set_column(1, 1, 15)
        sheet2.set_column(2, 2, 15)
        sheet2.set_column(3, 3, 15)
        sheet2.set_column(4, 4, 40)
        sheet2.set_column(5, 5, 40)
        sheet2.set_column(6, 6, 40)
        sheet2.set_column(7, 7, 15)
        workbook.close()
        return self.OutputFile

    def cleanup(self):
        if os.path.isfile(self.OutputFile) and os.stat(self.OutputFile).st_size > 0:
            try:
                os.remove(self.filepath)
            except:
                logger.debug(sys.exec_info()[0])
        else:
            raise Exception("Fehler beim bearbeiten der Datei: "+self.filepath.split("/")[-1])

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