import openpyxl
import time
import Tkinter as tk
import tkMessageBox
import pyttsx
import argparse
import sys
import os

def restartProgram():
    python = sys.executable
    os.execl(python, python, * sys.argv)
def getWorksheetAverages(worksheet):
    try:
        wb = openpyxl.load_workbook("SOBm-data.xlsx")
        sheet = wb.get_sheet_by_name(worksheet)
        avrHap = 0
        avrEne = 0
        avrFoc = 0
        for char in xrange(67, 70):
            avr = 0.0
            for row in xrange(2, sheet.max_row):
                if  sheet[str(chr(char)) + str(row)].value == None:
                    break
                #print sheet[str(chr(char)) + str(row)].value
                avr += sheet[str(chr(char)) + str(row)].value
            if char == 67:
                avrEne = avr / (sheet.max_row - 1)
            elif char == 68: 
                avrHap = avr / (sheet.max_row -1)
            else:
                avrFoc = avr / (sheet.max_row -1)
        print "Average happiness: {:.2f}\nAverage energy: {:.2f}\nAverage focus: {:.2f}".format(avrEne, avrHap, avrFoc)
                                                                               
    except IOError:
        print "No such file or directory: {}".format(filename)
        sys.exit()
    except KeyError:
        print "Worksheet '{}' does not exist.".format(worksheet)
        sys.exit()

def createSpreadsheet():
    try:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "1"
        sheet["A1"] = "Date"
        sheet["B1"] = "Time"
        sheet["C1"] = "Energy"
        sheet["D1"] = "Happiness"
        sheet["E1"] = "Focus"
        sheet["G1"] = "Spreadsheet_Line_To_Write"
        sheet["G2"] = 2
        wb.save("SOBm-data.xlsx")
    except:
        print "Something went wrong!\nTry again!"
    
def popupMessage(message): #show popup message
    root = tk.Tk()
    root.withdraw()
    tkMessageBox.showwarning('Self Stats', message)

def voiceMessage(text): #voice message
    engine = pyttsx.init()
    engine.say(text)
    engine.runAndWait()

def user_input(): #take user_input of happiness, focus and energy
    while True:
        
        try:
            happiness = int(raw_input("Enter happiness level between 1 and 10: ")) #happiness input
            if happiness < 1 or happiness > 10: #check if happiness is in range 1 to 10
                print "Sorry, wrong input! Try again."
                continue
            energy = int(raw_input("Enter energy level between 1 and 10: ")) #energy input
            if energy < 1 or energy > 10: #check if energy is in range 1 to 10
               print "Sorry, wrong input! Try again."
               continue
            focus = int(raw_input("Enter focus level between 1 and 10: ")) #focus input
            if focus < 1 or focus > 10: #check if focus is is range 1 to 10
               print "Sorry, wrong input! Try again."
               continue
        except: #in case of exception re-enter all values
            print "Something went wrong! Try again!"
            continue
        
        if 0 < focus <=10 and 0 < happiness <= 10 and 0 < energy <= 10: #when all values are between 1 and 10 brake the loop
            print "Input successful!"
            return [happiness, energy, focus] #return values as a list
            break

def writeToSpreadsheet(sheet, happiness, energy, focus, spreadsheet_line_to_write):
    try:
        print sheet
        pos = "A" + str(spreadsheet_line_to_write) #get spreadsheet_line_to_write eg. A2, from column name + spreadsheet_line_to_write (line number)
        sheet[pos] = time.strftime("%d.%m.%Y") #write date (dd.mm.yyyy) to sheet on spreadsheet_line_to_write eg. A2

        pos = "B" + str(spreadsheet_line_to_write) #get spreadsheet_line_to_write eg. B2, from column name + spreadsheet_line_to_write (line number)
        sheet[pos] = time.strftime("%H:%M:%S") #write time (hh:mm:ss) to sheet on spreadsheet_line_to_write eg. B2

        pos = "C" + str(spreadsheet_line_to_write) #get spreadsheet_line_to_write eg. C2, from column name + spreadsheet_line_to_write (line number)
        sheet[pos] = energy #write energy value to sheet on spreadsheet_line_to_write eg. C2

        pos = "D" + str(spreadsheet_line_to_write) #get spreadsheet_line_to_write eg. D2, from column name + spreadsheet_line_to_write (line number)
        sheet[pos] = happiness #write happiness value to sheet on spreadsheet_line_to_write eg. D2

        pos = "E" + str(spreadsheet_line_to_write) #get spreadsheet_line_to_write eg. E2, from column name + spreadsheet_line_to_write (line number)
        sheet[pos] = focus #write focus value to sheet on spreadsheet_line_to_write eg. E2

        spreadsheet_line_to_write += 1 #increase value of spreadsheet_line_to_write (sheet line) by one
        sheet['G2'] = spreadsheet_line_to_write #write spreadsheet_line_to_write value to the shell on spreadsheet_line_to_write G2

        print "Data successfuly entered into a sheet!"
        
    except:
        print "Something went wrong when entering data into a sheet!\nTry again!"
        sys.exit() 
        #restartProgram()

def saveSpreadsheet(wb):
    try:
        wb.save("SOBm-data.xlsx") #save spreadsheet to /home/bostjan/...
        print "File saved successfully!"

    except:
        print "Something went wrong while saving a file! File not saved!/nTry again!"
        
def wait(): #defines time between inputs
    print time.strftime("%H:%M:%S")
    print "Waiting 10 seconds."
    time.sleep(10)

def arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("-c", "--create-spreadsheet", action="store_true",
                        help="Create new spreadsheet in current folder.")
    parser.add_argument('-a', '--show-worksheet-averages', type=str,
                        help="Show averages for selected spreadsheet.")
    args = parser.parse_args()

    if args.create_spreadsheet == True:
        createSpreadsheet()
        sys.exit()

    if args.show_worksheet_averages is not None:
        arg = args.show_worksheet_averages
        getWorksheetAverages(arg)
        sys.exit()

def main():

    arguments()
    while True:
        wb = openpyxl.load_workbook("SOBm-data.xlsx")
        sheet = wb.get_sheet_by_name("May")
        spreadsheet_line_to_write = sheet['G2'].value
        voiceMessage("It's time to enter your feelings.")
        popupMessage("It's time to monitor your feelings")
        values = user_input()
        print values
        writeToSpreadsheet(sheet, values[0], values[1], values[2], spreadsheet_line_to_write)
        saveSpreadsheet(wb)
        wait()
        
if __name__=="__main__":
    main()
