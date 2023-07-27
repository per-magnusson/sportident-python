# -*- coding: utf-8 -*-
#
# Script to check if SI-cards listed in check.tsv has 
# punched in the controls listed in the same file.
# Each line in the file "check.tsv" has the format:
#
# SI-card_number<tab>electronic_control_number
#
# E.g.
#
# 721120	43
#
# All the files ending in .csv in the same directory are searched.
# The format of the CSV files is assumed to be that which is saved
# by Sportident Config+ when reading out the backup memory of 
# SI stations and saving via the function "Save current view" 
# (or something along those lines).
#
# In case not all control codes mentioned in check.tsv are found 
# in the csv files, a list of the missing control codes is printed.
#
# This program was hacked together by Per Magnusson, Linköpings OK,
# on 2019-07-25 and improved on 2019-07-27.
#
# The program is public domain and comes without any warranty. Feel free
# to do whatever you want with it, but don't blame me for any negative
# consequences.

import csv
import os
import time
from stat import *

CODE_FILE_NAME = 'check.tsv'
# Cards to be checked will be read into check_list
check_list = []
controls = set()

def delete_tsv_file(delete=True):
    if delete:
        try:
            os.remove(CODE_FILE_NAME)
        except Exception:
            print("Unknown error when deleting file")

def print_menu():
    print("** Current contents of %s **" % CODE_FILE_NAME)
    print("Enter SI-card and control number separated by space")
    print("Press enter to continue to log file check")
    print("Enter an index to delete the corresponding entry in the table")
    print("Enter 'p' to print a sorted list of all controls")
    print("Enter 'q' to quit")

def append_tsv_file():
    done_writing_file = False
    while not done_writing_file:
        try:
            with open(CODE_FILE_NAME, newline='') as tsvfile:
                temp_card_list = []
                list_reader = csv.reader(tsvfile, delimiter='\t')
                for row in list_reader:
                    temp_card_list.append([row[0], row[1]])
                # Print out current list of cards and controls
                print_menu()
                if len(temp_card_list) == 0:
                    print("Empty file %s" % CODE_FILE_NAME)
                else:
                    print("Index\tSI-Card\tControl")
                    for index, line in enumerate(temp_card_list):
                        print(str(index) + '\t' + line[0] + '\t' + line[1])
        except FileNotFoundError:
            # No file created yet, just print out instrutions on how to build the file
            print_menu()
            print("Empty file %s" % CODE_FILE_NAME)
        cmd = input("\"SI-card\" \"Control\": ")
        # Quit
        if cmd == 'q':
            exit()
        # Continue, finished editing tsv file
        elif cmd == '':
            done_writing_file = True
            continue
        # Print sorted list of all controls
        elif cmd == 'p':
            print("** All controls sorted by control code **")
            with open(CODE_FILE_NAME, newline='') as tsvfile:
                temp_control_list = []
                list_reader = csv.reader(tsvfile, delimiter='\t')
                for row in list_reader:
                    temp_control_list.append(int(row[1]))
                sort_temp_control_list = sorted(temp_control_list)
                for control in sort_temp_control_list:
                    print(control)
            exit()
        split_cmd = cmd.split(' ')
        # Add entry to tsv file
        if len(split_cmd) == 2:
            with open(CODE_FILE_NAME, 'a', newline='') as tsvfile:
                tsvfile.write(split_cmd[0] + '\t' + split_cmd[1] + '\n')
            continue
        # Delete entry in tsv file
        elif len(split_cmd) == 1:
            if split_cmd[0].isdigit():
                # Traverse the lines in the file, delete the input line number
                with open(CODE_FILE_NAME, 'r') as filedata:
                    # Read the file lines using readlines()
                    inputFilelines = filedata.readlines()
                    currentlineindex = 0
                    # Enter the line number to be deleted
                    deleteLine = int(split_cmd[0])
                    # Opening the given file in write mode.
                    with open(CODE_FILE_NAME, 'w') as filedata:
                        # Traverse in each line of the file
                        for textline in inputFilelines:
                            if currentlineindex != deleteLine:
                                # Write the corresponding line into file
                                filedata.write(textline)
                            # Increase the value of line index(line number) value by 1
                            currentlineindex += 1
            else:
                print(split_cmd, "is not a digit")
        else:
            print("Error: Could not parse input", split_cmd)
        

def check_and_read_tsv_file():
    # Try reading tsv file with cards to check into list of lists "check_list"
    try:
        with open(CODE_FILE_NAME, newline='') as tsvfile:
            list_reader = csv.reader(tsvfile, delimiter='\t')
            for row in list_reader:
                si_card = row[0]
                control = int(row[1])
                controls.add(control)
                check_list.append([si_card, control])
    except FileNotFoundError:
        print("Error: File not found")
        print("Start building new tsv file? y/n")
        input_char = input()
        if input_char != 'y':
            exit()
    except IndexError:
        print("IndexError when reading tsv file, is every line properly formatted?")
        print("Delete file and start building new tsv file? y/n")
        input_char = input()
        if input_char == 'y':
            delete_tsv_file()
    except Exception:
        print("Unknown error when reading tsv file")
        exit()

# First make sure the selected tsv file exists and is properly formatted
check_and_read_tsv_file()
# Next, prompt the user to add more cards and missing controls
append_tsv_file()
# Update the correct variables again now that the file has possibly been modified
check_and_read_tsv_file()

# Done creating input list, now start checking the read backup data
print("***** Start checking backup log data *****")

# Read all the .csv files in the same directory. 
# Sort on creation time, oldest first.

# Find all files in current dir and retrieve stats
file_names = (fn for fn in os.listdir('.'))
file_info = ((os.stat(name), name) for name in file_names)

# Keep only regular files and insert timestamp
file_info = ((stat[ST_CTIME], name)
           for stat, name in file_info if S_ISREG(stat[ST_MODE]))


for date, file_name in sorted(file_info):
    if(file_name.endswith(".csv")):
        with open(file_name, newline='') as csvfile:
            si_reader = csv.reader(csvfile, delimiter=';')
            si_reader.__next__()
            for row in si_reader:
                si_card = row[2]
                control = int(row[6])
                controls.discard(control)
                for check in check_list:
                    check_card = check[0]
                    check_control = check[1]
                    if check_control == control and check_card == si_card:
                        punch_time = row[3]
                        print ("Match on card: " + '{:8}'.format(int(si_card)) + 
                               " control: " + '{:3}'.format(control) + 
                               " time: " + punch_time + 
                               " file: " + file_name)

if len(controls) > 0:
    print("Missing logs from the following controls:")
    for code in sorted(controls):
        print(code)


print("***** Script finished *****")
        


        
