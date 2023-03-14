#!/usr/bin/python3

import subprocess
import csv
import os, sys
from chardet import detect

WB="../2016.xlsx"
CSV="2016.csv"
ETL="2016-etl.csv"

def get_encoding_type(file):
    with open(file, 'rb') as f:
        rawdata = f.read()
    return detect(rawdata)['encoding']

try:
    process = subprocess.Popen("xlsx2csv {} {}".format(WB,CSV), shell=True)
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

# 2016 format:
# Flight.Code, Flight.Name, Place, Bow, Name, Split 1, Split 2, Split 3, Raw Time, Adjustments, Adjusted Time
# Translation:
# Flight.Code [0] -> Event #
# Flight.Name has to be split as it contains the heat in the name
#     Men's Collegiate Novice Heat A
#     Women's Collegiate Novice Prelim
#     Freedom Rows Exhibition (1000m)
#     Men's Masters F
#     Women's Masters 8+F
#     Men's Collegiate Varsity 3rd Final
#     Women's Collegiate Varsity 3rd Final
#     Men's Alumni Petite Final
#     Women's Alumni Final Only
#     Men's Alumni Final Only
#     Men's Alumni Grand Final
#     Women's Collegiate Varsity DII/DIII/Club Petite Final
#     Women's Collegiate DII/DIII/Club 4+ Final Only

# Output:
# Event #, Trophy, Event, Heat, Place, Organization, Time


with open(CSV, 'r', encoding = get_encoding_type(CSV), errors='ignore') as infile:
    inreader = csv.reader(infile, dialect='excel')

    for row in inreader:
        rowcount += 1

        if(rowcount < 1 or rowcount >696):
            continue

        if(rowcount == 1):
            row = ["Event #", "Trophy", "Event", "Heat", "Place", "Boat", "Time", "Cox", "Rower1", "Rower2", "Rower3", "Rower4", "Rower5", "Rower6", "Rower7", "Rower8"]
            outwriter.writerow(row)
        else:
            newrow = []
            for i in range(0,18):
                newrow.append('')
            newrow[0] = row[0] # Event #
            newrow[1] = ''   # 2016 data contains no trophy information
            flightname = row[1]
            if "Heat " in flightname:
                (title, heat) = flightname.split("Heat")
            elif "Prelim" in flightname:
                title = flightname.replace("Prelim","")
                heat = "Prelim"
            elif "3rd Final" in flightname:
                title = flightname.replace("3rd Final","")
                heat = "3rd Final"
            elif "Petite Final" in flightname:
                title = flightname.replace("Petite Final","")
                heat = "Petite Final"
            elif "Final Only" in flightname:
                title = flightname.replace("Final Only","")
                heat = "Final Only"
            elif "Grand Final" in flightname:
                title = flightname.replace("Grand Final","")
                heat = "Grand Final"
            else:
                title = flightname
                heat = ""
                
            newrow[2] = title.strip() # Event title
            newrow[3] = heat.strip()  # Heat
            newrow[4] = row[2] # place
            newrow[5] = row[4] # boat
            newrow[6] = row[8] # time

            print(newrow)
            outwriter.writerow(newrow)


infile.close()
outfile.close()
os.unlink(CSV)
