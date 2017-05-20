import sys
import os
import mysql.connector
import time
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

def getAverages(master):
    try:
        con = mysql.connector.connect(user='sob', password=master, host='localhost',
                                  database='sob_monitor')
        cursor = con.cursor()
        get_number_of_rows = "SELECT COUNT(*) FROM sob_data;"
        get_happiness = "SELECT happiness FROM sob_data;"
        get_energy = "SELECT energy FROM sob_data;"
        get_focus = "SELECT focus FROM sob_data;"
        cursor.execute(get_number_of_rows)
        for value in cursor:
            number_of_rows = value[0]

        cursor.execute(get_happiness)
        happiness = 0
        for value in cursor:
            happiness += value[0]
        average_happiness = happiness / number_of_rows

        cursor.execute(get_energy)
        energy = 0
        for value in cursor:
            energy += value[0]
        average_energy = energy / number_of_rows

        cursor.execute(get_focus)
        focus = 0
        for value in cursor:
            focus += value[0]
        average_focus = focus / number_of_rows
        con.close()
        print "Average happiness: {:.2f}\nAverage energy: {:.2f}\nAverage focus: {:.2f}".format(average_energy, average_happiness, average_focus)

    except Exception, e:
        print Bcolors.FAIL + repr(e) + Bcolors.ENDC
        print Bcolors.WARNING + "Program exited!" + Bcolors.ENDC
        sys.exit()

def getAveragesInRange():
    try:
        start = raw_input("Enter start date (yyyy-mm-dd): ")
        end = raw_input("Enter end date (yyyy-mm-dd): ")
        host = raw_input("Host: ")
        passwd = getpass("Password for user 'sob': ")
        con = mysql.connector.connect(user="sob", host=host, password=passwd, database="sob_monitor")
        cursor = con.cursor()
        number_of_rows = "SELECT COUNT(*) FROM sob_data WHERE date(date_date) between '{}' and '{}';".format(start, end)
        get_happiness = "SELECT happiness FROM sob_data WHERE date(date_date) between '{}' and '{}';".format(start, end)
        get_focus = "SELECT focus FROM sob_data WHERE date(date_date) between '{}' and '{}';".format(start, end)
        get_energy = "SELECT energy FROM sob_data WHERE date(date_date) between '{}' and '{}';".format(start, end)

        cursor.execute(number_of_rows)
        number_of_rows = cursor.fetchall()
        number_of_rows = int(number_of_rows[0][0])

        cursor.execute(get_happiness)
        happiness = 0
        for value in cursor:
            happiness += value[0]
        average_happiness = happiness / number_of_rows

        cursor.execute(get_energy)
        energy = 0
        for value in cursor:
            energy += value[0]
        average_energy = energy / number_of_rows

        cursor.execute(get_focus)
        focus = 0
        for value in cursor:
            focus += value[0]
        average_focus = focus / number_of_rows
        con.close()
        print "Average happiness: {:.2f}\nAverage energy: {:.2f}\nAverage focus: {:.2f}".format(average_energy, average_happiness, average_focus)

    except IOError:
        print Bcolors.WARNING + "No such file or directory: {}\nProgram exited!".format(filename) + Bcolors.ENDC
        sys.exit()
    except KeyError:
        print Bcolors.WARNING + "Worksheet '{}' does not exist.\nProgram exited!".format(worksheet) + Bcolors.ENDC
        sys.exit()
    except Exception, e:
        print Bcolors.FAIL + repr(e) + Bcolors.ENDC
        print Bcolors.WARNING + "Program exited!" + Bcolors.ENDC
        sys.exit()

def createDatabase():
    try:
        print Bcolors.OKBLUE + "You are in database creation mode." + Bcolors.ENDC
        host = raw_input("Host: ")
        user = raw_input("User (with full privileges): ")
        passwd = getpass("Password for '{}': ".format(user))
        con = mysql.connector.connect(user=user, host=host, password=passwd)
        cursor = con.cursor()
        cursor.execute("show databases;")
        data = cursor.fetchall()
        for value in data:
            if value[0] == "sob_monitor":
                print "{}Database {} already exist!\nProgram exited!{}".format(Bcolors.WARNING, value[0], Bcolors.ENDC)
                sys.exit()
        cursor.execute("CREATE DATABASE sob_monitor;")
        cursor.execute("USE sob_monitor;")
        cursor.execute("Create table sob_data (entry_id Int NOT NULL AUTO_INCREMENT, date_date Date NOT NULL, date_time Time NOT NULL, happiness Int NOT NULL, energy Int NOT NULL, focus Int NOT NULL, UNIQUE (entry_id), Index AI_entry_id (entry_id), Primary Key (entry_id)) ENGINE = MyISAM;")
        cursor.execute("Create table passwd (entry_id Int NOT NULL AUTO_INCREMENT, hash Varchar(150) NOT NULL, salt Varchar(150) NOT NULL, UNIQUE (entry_id), UNIQUE (hash), Index AI_entry_id (entry_id), Primary Key (entry_id)) ENGINE = MyISAM;")

    except Exception, e:
        print Bcolors.FAIL + repr(e) + Bcolors.ENDC
        print Bcolors.WARNING + "Program exited!" + Bcolors.ENDC
        sys.exit()
        
    try:
        print "{}Creating user 'sob'...{}".format(Bcolors.OKBLUE, Bcolors.ENDC)
        cursor.execute("drop user 'sob'@'{}'".format(host))
        cursor.execute("flush privileges;")
        passwd = getpass("Choose password for user 'sob': ")
        cursor.execute("Create user 'sob'@'{}' identified by '{}'".format(host, passwd))
        cursor.execute("GRANT insert, select on sob_monitor.* to 'sob'@'{}'".format(host))
        print "{}User 'sob' successfully crated!{}".format(Bcolors.OKGREEN, Bcolors.ENDC)

    except Exception, e:
        print Bcolors.FAIL + repr(e) + Bcolors.ENDC
        print Bcolors.WARNING + "Program exited!" + Bcolors.ENDC
        sys.exit()

    try:
        print "{}Creating user 'sob_data' (without password)...!{}".format(Bcolors.OKBLUE, Bcolors.ENDC)
        cursor.execute("drop user 'sob_data'@'{}'".format(host))
        cursor.execute("flush privileges;")
        cursor.execute("Create user 'sob_data'@'{}'".format(host))
        cursor.execute("GRANT insert on sob_monitor.sob_data to 'sob_data'@'{}'".format(host))
        cursor.execute("GRANT select on sob_monitor.passwd to 'sob_data'@'{}'".format(host))
        print "{}User 'sob_data' successfully crated!{}".format(Bcolors.OKGREEN, Bcolors.ENDC)

    except Exception, e:
        print Bcolors.FAIL + repr(e) + Bcolors.ENDC
        print Bcolors.WARNING + "Program exited!" + Bcolors.ENDC
        sys.exit()

def write(date, hour, happiness, energy, focus):
    con = mysql.connector.connect(user='sob_data', password='', host='localhost', database='sob_monitor')
    cursor = con.cursor()
    data = "INSERT INTO sob_data VALUES({}, {}, {}, {}, {}, {});".format(0, date, hour, happiness, energy, focus)
    cursor.execute(data)
    con.close()

def main(values):
    try:
        hour = time.strftime("%H%M%S")
        date = time.strftime("%Y%m%d")
        #print values
        write(date, hour, values[0], values[1], values[2])
        print Bcolors.OKGREEN + "At " + time.strftime("%H:%M:%S,") +  " data was successfully written to the database!" + Bcolors.ENDC 

    except Exception, e:
        print Bcolors.WARNING + "Failed writing data to database!" + Bcolors.ENDC
        print Bcolors.FAIL + repr(e) + Bcolors.ENDC
        print Bcolors.WARNING + "Program exited!" + Bcolors.ENDC
        sys.exit()
