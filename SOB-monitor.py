import openpyxl
import time
import Tkinter as tk
import tkMessageBox
import pyttsx
import argparse
import sys
import os
import mysql.connector
import bcrypt
import db_mode
import spreadsheet_mode
from getpass import getpass

class Bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def restartProgram():
    python = sys.executable
    os.execl(python, python, * sys.argv)
    
def popupMessage(message): #show popup message
    root = tk.Tk()
    root.withdraw()
    tkMessageBox.showwarning('SOB-monitor', message)

def voiceMessage(text): #voice message
    engine = pyttsx.init()
    engine.say(text)
    engine.runAndWait()

def user_input(): #take user_input of happiness, focus and energy
    while True:
        
        try:
            happiness = int(raw_input("Enter happiness level between 1 and 10: ")) #happiness input
            if happiness < 1 or happiness > 10: #check if happiness is in range 1 to 10
                print Bcolors.WARNING + "Sorry, wrong input! Try again." + Bcolors.ENDC
                continue
            energy = int(raw_input("Enter energy level between 1 and 10: ")) #energy input
            if energy < 1 or energy > 10: #check if energy is in range 1 to 10
               print Bcolors.WARNING + "Sorry, wrong input! Try again." + Bcolors.ENDC
               continue
            focus = int(raw_input("Enter focus level between 1 and 10: ")) #focus input
            if focus < 1 or focus > 10: #check if focus is is range 1 to 10
               print Bcolors.WARNING + "Sorry, wrong input! Try again." + Bcolors.ENDC
               continue

            if 0 < focus <=10 and 0 < happiness <= 10 and 0 < energy <= 10: #when all values are between 1 and 10 brake the loop
                #print Bcolors.OKGREEN + "Input successful!" + Bcolors.ENDC
                return [happiness, energy, focus] #return values as a list
                break

        except Exception, e:
            print Bcolors.FAIL + repr(e) + Bcolors.ENDC
            print Bcolors.WARNING + "Program exited!" + Bcolors.ENDC
            sys.exit()
        
def wait(wait_time): #defines time between inputs
    #print time.strftime("%H:%M:%S")
    print "{} seconds till next entry.".format(wait_time)
    time.sleep(wait_time)

def arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("-c", "--create-spreadsheet", action="store_true",
                        help="Create new spreadsheet in current folder.")
    parser.add_argument('-a', '--show-averages', action="store_true",
                        help="Show averages of all entries.")
    parser.add_argument('--database-mode', action="store_true",
                        help="Use database (you need to create it) instead of spreadsheet.")
    parser.add_argument("--sheet-name", type=str, help="Name of the sheet in wich data will be saved. Default: 'data'")
    parser.add_argument("--create-database", action="store_true", help="Crate database for storing data.")
    parser.add_argument("--choose-range", action="store_true", help="Choose range in wich averages will be calculated.")
    parser.add_argument("--wait", type=int, help="Set time (in seconds) between entries.")
    args = parser.parse_args()

    if args.database_mode == True:
        global mode
        mode = "database"
    else:
        mode = "spreadsheet"

    if args.create_spreadsheet == True:
        spreadsheet_mode.createSpreadsheet()
        sys.exit()

    if args.show_averages == True and mode == "database" and args.choose_range != True:
        arg = args.show_averages
        master = getpass("Password for 'scooter':" )
        db_mode.getAverages(master)
        sys.exit()

    elif args.show_averages == True and mode == "spreadsheet" and args.choose_range != True:
        sheet = raw_input("Enter the name of sheet: ")
        spreadsheet_mode.getAverages(sheet)
        sys.exit()

    if args.sheet_name is not None:
        global sheet_name
        sheet_name = args.sheet_name
    else:
        sheet_name = "data"

    if args.create_database == True:
        db_mode.createDatabase()
        sys.exit()

    if args.choose_range == True and args.show_averages == True:
        spreadsheet_mode.getAveragesInRange()
        sys.exit()

    if args.wait is not None:
        global wait_time
        wait_time = args.wait
    else:
        wait_time = 3600

def main():
    arguments()
    
    while True:
        voiceMessage("It's time to enter your feelings.")
        popupMessage("It's time to enter your feelings!")
        values = user_input()
        if mode == "spreadsheet":
            spreadsheet_mode.main(values, sheet_name)
        else:
            db_mode.main(values)
        wait(wait_time)
        
if __name__=="__main__":
    main()
