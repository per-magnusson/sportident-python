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


# Read tsv file with cards to check into list of lists "check_list"
check_list = []
controls = set()
with open('check.tsv', newline='') as tsvfile:
    list_reader = csv.reader(tsvfile, delimiter='\t')
    for row in list_reader:
        si_card = row[0]
        control = int(row[1])
        controls.add(control)
        check_list.append([si_card, control])

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
        


        
