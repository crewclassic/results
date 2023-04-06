#!/usr/bin/python3

import subprocess
import csv
import os, sys
from chardet import detect

YEAR="1973_through_2003"
WB="../{}.xls".format(YEAR)
CSV="{}.csv".format(YEAR)
ETL="{}-etl.csv".format(YEAR)

def get_encoding_type(file):
    with open(file, 'rb') as f:
        rawdata = f.read()
    return detect(rawdata)['encoding']

try:
    process = subprocess.Popen("/usr/bin/libreoffice --headless --convert-to csv {} -o .".format(WB), shell=True)
    process.wait()
    
except:
    print("Unable to convert {} to csv".format(WB))
    exit()

outfile = open(ETL,'w')
outwriter = csv.writer(outfile,dialect='excel', lineterminator='\n')

rowcount=0
trophy = ""
current_trophy = ""
event = ""

# 1973-2003 format
# header on line 0
# SortOrder, Year, Place, EventType, Event, Team, Time, Cox, Rower1, Rower2, Rower3, Rower4, Rower5, Rower6, Rower7, Rower8
#
# EventType was originally not used, then Petite Final was put in
# there and in the late 1990's they started putting the Trophy name
# into the column. In this ET process, we will use EventType as Trophy
# and exclude any mention of "Petite Final"


# Output:
# Year, Event #, Trophy, Event, Heat, Place, Boat, Time, Cox, Rower1, Rower2, Rower3, Rower4, Rower5, Rower6, Rower7, Rower8


with open(CSV, 'r', encoding = get_encoding_type(CSV), errors='ignore') as infile:
    inreader = csv.reader(infile, dialect='excel')

    row = ["Year", "Event #", "Trophy", "Event", "Heat", "Place", "Boat", "Time", "Cox", "Rower1", "Rower2", "Rower3", "Rower4", "Rower5", "Rower6", "Rower7", "Rower8"]
    outwriter.writerow(row)

    state = 0
    place = 1
    last_event_number = 0
    for row in inreader:
        rowcount += 1
        if state == 0:
            if row[0] == "SortOrder":
                state = 1
        elif state == 1:
            if row[0] == '':
                state = 2
            else:
                newrow = []
                for i in range(0,17):
                    newrow.append('')

                newrow[0] = row[1] # Year
                newrow[1] = ''     # Event #'s are not recorded in this data
                if row[3].lower() != 'petite final':
                    newrow[2] = row[3]   # Trophy from EventType column

                newrow[3] = row[4] # Event title
                newrow[4] = ''     # Heat's were not recorded as a separate column
                newrow[5] = row[2] # Place
                newrow[6] = row[5] # Boat
                newrow[7] = row[6] # Time
                for i in range(8,17):
                    newrow[i] = row[i-1]

                outwriter.writerow(newrow)
        elif state == 2:
            pass
        
infile.close()
outfile.close()
os.unlink(CSV)
