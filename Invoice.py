import xmltodict
import xlsxwriter
import os
from json import loads, dumps
import logging

class Invoice:

    def __init__(self, file, config):
        self.filepath = config["files"]["source_path"]+file
        with open(self.filepath, "r", encoding="utf-8") as fd:
            self.config = config
            self.doc = xmltodict.parse(fd.read())
            self.FileSender = self.doc["epay21Finance"]["PSPData"]["FileSender"]
            self.FileName = self.doc["epay21Finance"]["PSPData"]["FileName"]
            # self.OutputFile = config["files"]["destination_path"]+self.FileName.split(".")[0]+".xlsx"
            self.OutputFile = config["files"]["destination_path"]+file.replace(".xml",".xlsx")
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
        usk_config = loads(self.config["usk"]["usk_liste"].replace("\'","\""))
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
