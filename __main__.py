from tkinter import filedialog
from tkinter import *
import re
import codecs
from xlrd import open_workbook
from xlutils.copy import copy
from datetime import *
import ctypes
import os


root = Tk()


def main():
    cwd = os.getcwd()
    MessageBox = ctypes.windll.user32.MessageBoxW
    MessageBox(
        None,
        "Choose your Scholarshipowl.com HTML file.\nTo get this file, use CTRL+S while visiting ScholarshipOwl.com",
        "ScholarshipOwl Export",
        0,
    )
    HTMLFilename = filedialog.askopenfilename(
        initialdir=cwd,
        title="Select scholarshipowl.com .HTML File",
        filetypes=(("html files", "*.html"), ("all files", "*.*")),
    )
    MessageBox(
        None,
        "Choose your Scholarshipowl.com Excel .XLS File.",
        "ScholarshipOwl Export",
        0,
    )
    excelFileName = filedialog.askopenfilename(
        initialdir=cwd,
        title="Select Scholarship .XLS Document",
        filetypes=((".xls Files", "*.xls"), ("all files", "*.*")),
    )
    oldScholarshipsInfo = loadScholarshipsFromExcel(excelFileName)
    oldScholarships = oldScholarshipsInfo[0]
    cellValueIndex = oldScholarshipsInfo[1]
    newScholarships = loadScholarshipsFromHTML(HTMLFilename)
    scholarshipsToAdd = findNewScholarships(oldScholarships, newScholarships)
    updateExcel(excelFileName, scholarshipsToAdd, cellValueIndex)


def updateExcel(excelDocumentLocation, scholarshipsToAdd, cellValueIndex):
    saveTriesLeft = 3
    print("updating excel")
    while saveTriesLeft >= 0:
        try:
            rb = open_workbook(excelDocumentLocation, formatting_info=True)
            wb = copy(
                rb
            )  # a writable copy (I can't read values out of this, only write to it)
            for scholarship in scholarshipsToAdd:
                print("Adding scholarship: {}".format(scholarship))
                columnPrint = 0
                for item in scholarship:
                    try:
                        if "http" not in item:
                            item = item.title()
                        wb.get_sheet(0).write(cellValueIndex, columnPrint, item)
                    except:
                        pass
                    columnPrint += 1
                cellValueIndex += 1
            wb.save(excelDocumentLocation)
            saveTriesLeft = -1
        except:
            MessageBox = ctypes.windll.user32.MessageBoxW
            MessageBox(
                None,
                "Close out of Scholarship Excel File (.xls). Trying to save {} more times.".format(
                    saveTriesLeft
                ),
                "Error: ScholarshipOwl Export",
                0,
            )
            saveTriesLeft -= 1


def findNewScholarships(oldScholarships, newScholarships):
    newScholarshipList = [
        scholarship
        for scholarship in newScholarships
        if scholarship not in oldScholarships
    ]
    return newScholarshipList


def loadScholarshipsFromHTML(htmlDocumentLocaiton):
    file = codecs.open(htmlDocumentLocaiton, "r", "utf-8")
    html = file.read()
    html = html.split("""/thead>""")[1].split("""/table>""")[0].split("</tbody>", 1)[0]
    scholarships = html.split(
        """class="mod-td-checkbox"><input name="scholarship_id[]" type="hidden" value="""
    )
    scholarships.pop(0)
    scholarshipList = []
    for scholarship in scholarships:
        scholarshipName = (
            scholarship.split("<strong>", 1)[1]
            .split("</strong>")[0]
            .split('<i class="icon icon', 1)[0]
            .rstrip()
            .lstrip()
        )
        scholarshipLink = scholarship.split('href="', 1)[1].split('" target=', 1)[0]
        deadlineBlock = scholarship.split("deadline", 1)[1].split("amount", 1)[0]
        scholarshipDate = findDate(deadlineBlock)
        scholarshipAmount = scholarship.split(""""mod-td-ammount"><strong>""", 1)[
            1
        ].split("</strong>", 1)[0]
        if "recurrent" in scholarship:
            scholarshipRecurrent = "Yes"
        else:
            scholarshipRecurrent = "No"
        scholarshipInfo = [
            scholarshipName.lower().lstrip().rstrip(),
            scholarshipAmount.lower().lstrip().rstrip(),
            scholarshipDate.lower().lstrip().rstrip(),
            scholarshipLink.lower().lstrip().rstrip(),
            scholarshipRecurrent.lower().lstrip().rstrip(),
        ]
        scholarshipList.append(scholarshipInfo)
    return scholarshipList


def loadScholarshipsFromExcel(excelDocumentLocation):
    rb = open_workbook(excelDocumentLocation, formatting_info=True)
    r_sheet = rb.sheet_by_index(0)  # read only copy to introspect the file
    scholarshipList = []
    cellValueIndex = 1
    scholarshipName = "null"
    scholarshipNameColumn = 0
    scholarshipAmountColumn = 1
    scholarshipDueDateColumn = 2
    scholarshipLinkColumn = 3
    scholarshipRecurringColumn = 4
    try:
        while scholarshipName != "":
            scholarshipName = r_sheet.cell(cellValueIndex, scholarshipNameColumn).value
            scholarshipAmount = r_sheet.cell(
                cellValueIndex, scholarshipAmountColumn
            ).value
            scholarshipLink = r_sheet.cell(cellValueIndex, scholarshipLinkColumn).value
            scholarshipDueDate = r_sheet.cell(
                cellValueIndex, scholarshipDueDateColumn
            ).value
            scholarshipRecurring = r_sheet.cell(
                cellValueIndex, scholarshipRecurringColumn
            ).value
            scholarshipList.append(
                [
                    scholarshipName.lower().lstrip().rstrip(),
                    scholarshipAmount,
                    scholarshipDueDate.lower().lstrip().rstrip(),
                    scholarshipLink.lower().lstrip().rstrip(),
                    scholarshipRecurring.lower().lstrip().rstrip(),
                ]
            )
            cellValueIndex += 1
            scholarshipName = r_sheet.cell(cellValueIndex, 0).value
    except:
        print("test")
        pass
    return [scholarshipList, cellValueIndex]


def findDate(inputString):
    dateRegexPrecise = r"(?i)(Jan(uary)?|Feb(ruary)?|Mar(ch)?|Apr(il)?|May|June?|July?|Aug(ust)?|Sep(tember)?|Oct(ober)?|Nov(ember)?|Dec(ember)?).?,?\s(\d\d?)([^\d]{0,5})(\d\d\d\d)"
    dateRegexGeneral = r"(?i)(Jan(uary)?|Feb(ruary)?|Mar(ch)?|Apr(il)?|May|June?|July?|Aug(ust)?|Sep(tember)?|Oct(ober)?|Nov(ember)?|Dec(ember)?).?,?\s(\d\d?)([^\d]{0,5})(\d\d\d\d)?"
    dateRegexFormatted = (
        r"((0|1|2|3)?\d{1})(-|/)((0|1|2|3)\d{1})(-|/)(((19|20)(\d{2})?))"
    )
    dateRe = re.findall(dateRegexPrecise, inputString)
    if len(dateRe) == 0:
        dateRe = re.findall(dateRegexGeneral, inputString)
    if len(dateRe) == 0:
        dateRe = re.findall(dateRegexFormatted, inputString)
    if len(dateRe) > 0:
        dateRe = dateRe[0]
        if len(dateRe) == 10:
            month = dateRe[0]
            day = dateRe[3]
            year = dateRe[6]
        else:
            month = str(monthStrToInt(dateRe[0]))
            day = dateRe[10]
            year = dateRe[12]
            if year == "":
                year = str(datetime.now().year)
        dateStr = year + "/" + month + "/" + day
        return dateStr


def monthStrToInt(month):
    if month.lower().startswith("jan"):
        return 1
    elif month.lower().startswith("feb"):
        return 2
    elif month.lower().startswith("mar"):
        return 3
    elif month.lower().startswith("apr"):
        return 4
    elif month.lower().startswith("may"):
        return 5
    elif month.lower().startswith("jun"):
        return 6
    elif month.lower().startswith("jul"):
        return 7
    elif month.lower().startswith("aug"):
        return 8
    elif month.lower().startswith("sep"):
        return 9
    elif month.lower().startswith("oct"):
        return 10
    elif month.lower().startswith("nov"):
        return 11
    elif month.lower().startswith("dec"):
        return 12
    else:
        return 00


main()
