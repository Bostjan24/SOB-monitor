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
def getWorksheetAverages(filename, worksheet):
    try:
        wb = openpyxl.load_workbook(filename)
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
        print "Average happiness: %.2f \nAverage energy: %.2f \nAverage focus: %.2f" %(avrHap,
                                                                                       avrEne, avrFoc)
    except IOError:
        print "No such file or directory: %s" %filename
        sys.exit()
    except KeyError:
        print "Worksheet '%s' does not exist." %worksheet

def createSpreadsheet(filename):
    try:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "1"
        sheet["A1"] = "Date"
        sheet["B1"] = "Time"
        sheet["C1"] = "Energy"
        sheet["D1"] = "Happiness"
        sheet["E1"] = "Focus"
        sheet["G1"] = "Position"
        sheet["G2"] = 2
        wb.save(filename)
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

def input(): #take input of happiness, focus and energy
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

def writeToSpreadsheet(sheet, happiness, energy, focus, position):
    try:
        pos = "A" + str(position) #get position eg. A2, from column name + position (line number)
        sheet[pos] = time.strftime("%d.%m.%Y") #write date (dd.mm.yyyy) to sheet on position eg. A2

        pos = "B" + str(position) #get position eg. B2, from column name + position (line number)
        sheet[pos] = time.strftime("%H:%M:%S") #write time (hh:mm:ss) to sheet on position eg. B2

        pos = "C" + str(position) #get position eg. C2, from column name + position (line number)
        sheet[pos] = energy #write energy value to sheet on position eg. C2

        pos = "D" + str(position) #get position eg. D2, from column name + position (line number)
        sheet[pos] = happiness #write happiness value to sheet on position eg. D2

        pos = "E" + str(position) #get position eg. E2, from column name + position (line number)
        sheet[pos] = focus #write focus value to sheet on position eg. E2

        position += 1 #increase value of position (sheet line) by one
        sheet['G2'] = position #write position value to the shell on position G2

        print "Data seccessfuly entered into a sheet!"
        
    except:
        print "Something went wrong when entering data into a sheet!\nTry again!"
        sys.exit() 
        #restartProgram()

def saveSpreadsheet(wb):
    try:
        wb.save("/home/bostjan/Documents/sheetTest.xlsx") #save spreadsheet to /home/bostjan/...
        print "File saved successfully!"

    except:
        print "Something went wrong while saving a file! File not saved! '/n' Try again!"
        
def wait(): #defines time between inputs
    print time.strftime("%H:%M:%S")
    print "Waiting 10 seconds."
    time.sleep(10)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-c", "--create-spreadsheet",
                        help="Create new spreadsheet in current folder.")
    parser.add_argument('-a', '--show-worksheet-averages', nargs='+', type=str,
                        help="ged")
    args = parser.parse_args()

    if args.create_spreadsheet is not None:
        createSpreadsheet(args.create_spreadsheet)
        sys.exit()

    if args.show_worksheet_averages is not None:
        arg = args.show_worksheet_averages
        getWorksheetAverages(arg[0], arg[1])
        sys.exit()
    
    while True:
        wb = openpyxl.load_workbook("/home/bostjan/Documents/sheet.xlsx") #open spreadsheet from /home/bostjan/...
        sheet = wb.get_sheet_by_name("May") #select sheet named 'May'
        position = sheet['G2'].value #from sheet position G2 get value of position (line number)
        voiceMessage("It's time to enter your feelings.")
        popupMessage("It's time to monitor your feelings") #show popup message
        values = input() #input part
        writeToSpreadsheet(sheet, values[0], values[1], values[2], position) #write values to sheet
        saveSpreadsheet(wb) #save spreadsheet
        wait()
        
if __name__=="__main__":
    main()
