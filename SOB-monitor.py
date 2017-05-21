import time
import pyttsx
import Tkinter as tk
import tkMessageBox
import os
import sys
import argparse
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
    
def popupMessage(message):
    root = tk.Tk()
    root.withdraw()
    tkMessageBox.showwarning('SOB-monitor', message)

def voiceMessage(message):
    engine = pyttsx.init()
    engine.say(message)
    engine.runAndWait()

def user_input(): #takes user input of happiness, focus and energy
    while True:
        
        try:
            happiness = int(raw_input("Enter happiness level between 1 and 10: "))
            if happiness < 1 or happiness > 10:
                print Bcolors.WARNING + "Sorry, wrong input! Try again." + Bcolors.ENDC
                continue
            energy = int(raw_input("Enter energy level between 1 and 10: "))
            if energy < 1 or energy > 10:
               print Bcolors.WARNING + "Sorry, wrong input! Try again." + Bcolors.ENDC
               continue
            focus = int(raw_input("Enter focus level between 1 and 10: "))
            if focus < 1 or focus > 10:
               print Bcolors.WARNING + "Sorry, wrong input! Try again." + Bcolors.ENDC
               continue

            if 0 < focus <=10 and 0 < happiness <= 10 and 0 < energy <= 10: #when all values are between 1 and 10 brake the loop
                return [happiness, energy, focus]
                break

        except Exception, e:
            print Bcolors.FAIL + repr(e) + Bcolors.ENDC
            print Bcolors.WARNING + "Program exited!" + Bcolors.ENDC
            sys.exit()
        
def wait(wait_time):
    print "{} seconds till next entry.".format(wait_time)
    time.sleep(wait_time)

def arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("--create-spreadsheet", action="store_true",
                        help="Create new spreadsheet in current folder.")
    parser.add_argument('-a', '--show-averages', action="store_true",
                        help="Show averages of all entries.")
    parser.add_argument('--database-mode', action="store_true",
                        help="Use database instead of spreadsheet.")
    parser.add_argument("--sheet-name", type=str, help="Name of the sheet in wich data will be saved. Default: 'data'")
    parser.add_argument("--create-database", action="store_true", help="Crate database for storing data.")
    parser.add_argument("--range", action="store_true", help="Choose range in wich averages will be calculated.")
    parser.add_argument("--wait", type=int, help="Set time (in seconds) between entries.")
    args = parser.parse_args()

    if args.create_spreadsheet == True and args.database_mode == False:
        spreadsheet_mode.createSpreadsheet()
        sys.exit()

    if args.show_averages == True  and args.database_mode == True and args.choose_range != True:
        arg = args.show_averages
        master = getpass("Password for 'sob':" )
        averages = db_mode.getAverages(master)
        print "Average happiness: {:.2f}\nAverage energy: {:.2f}\nAverage focus: {:.2f}".format(averages[0], averages[1], averages[2])
        sys.exit()

    elif args.show_averages == True and args.database_mode == False and args.choose_range != True:
        sheet = raw_input("Enter the name of sheet: ")
        averages = spreadsheet_mode.getAverages(sheet)
        print "Average happiness: {:.2f}\nAverage energy: {:.2f}\nAverage focus: {:.2f}".format(averages[0], averages[1], averages[2])
        sys.exit()

    if args.sheet_name is not None:
        global sheet_name
        sheet_name = args.sheet_name
    else:
        sheet_name = "data"

    if args.create_database == True and args.database_mode == True:
        db_mode.createDatabase()
        sys.exit()

    if args.choose_range == True and args.show_averages == True and args.database_mode == False:
        averages = spreadsheet_mode.getAveragesInRange()
        print "Average happiness: {:.2f}\nAverage energy: {:.2f}\nAverage focus: {:.2f}".format(averages[0], averages[1], averages[2])
        sys.exit()

    if args.choose_range == True and args.show_averages == True and args.database_mode == True:
        averages = db_mode.getAveragesInRange()
        print "Average happiness: {:.2f}\nAverage energy: {:.2f}\nAverage focus: {:.2f}".format(averages[0], averages[1], averages[2])
        sys.exit()

    if args.wait is not None:
        global wait_time
        wait_time = args.wait
    else:
        wait_time = 3600

    if len(sys.argv) >= 1 and args.database_mode == False and args.create_database == True:
        print "{}You are not in database mode!\nProgram exited!{}".format(Bcolors.WARNING, Bcolors.ENDC)
        sys.exit()
    elif len(sys.argv) >= 1 and args.database_mode == True and args.create_spreadsheet == True:
        print "{}You are not in spreadsheet mode!\nProgram exited!{}".format(Bcolors.WARNING, Bcolors.ENDC)
        sys.exit()

    if args.database_mode == False:
        voiceMessage("It's time to enter your feelings.")
        popupMessage("It's time to enter your feelings!")
        values = user_input()
        spreadsheet_mode.main(values, sheet_name)
    elif args.database_mode == True:
        voiceMessage("It's time to enter your feelings.")
        popupMessage("It's time to enter your feelings!")
        values = user_input()
        db_mode.main(values)

def main():
    while True:
        arguments()
        wait(wait_time)
        
if __name__=="__main__":
    main()
