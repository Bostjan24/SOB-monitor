import openpyxl
import time
import Tkinter as tk
import tkMessageBox
import pyttsx
import os

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

def saveSpreadsheet(wb):
    wb.save("/home/bostjan/Documents/sheetTest.xlsx") #save spreadsheet to /home/bostjan/...

def wait(): #defines time between inputs
    print "Waiting 10 seconds."
    time.sleep(10)

def main():
    while True:
        wb = openpyxl.load_workbook("/home/bostjan/Documents/sheetTest.xlsx") #open spreadsheet from /home/bostjan/...
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