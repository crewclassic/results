#!/usr/bin/python3

import subprocess
import csv
import os, sys
from chardet import detect

YEAR="2007"
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

# 2005 format
# Event #, Start, Title, Heat, Trophy, 1st Place Organization, 1st Place time... 7th Place Organization, 7th Place time
# Output:
# Year, Event #, Trophy, Event, Heat, Place, Boat, Time, Cox, Rower1, Rower2, Rower3, Rower4, Rower5, Rower6, Rower7, Rower8


with open(CSV, 'r', encoding = get_encoding_type(CSV), errors='ignore') as infile:
    inreader = csv.reader(infile, dialect='excel')

    for row in inreader:
        rowcount += 1

        if(rowcount >= 58 and rowcount <= 61 or rowcount < 2):
            continue

        if(rowcount == 2):
            row = ["Year", "Event #", "Trophy", "Event", "Heat", "Place", "Boat", "Time", "Cox", "Rower1", "Rower2", "Rower3", "Rower4", "Rower5", "Rower6", "Rower7", "Rower8"]
            outwriter.writerow(row)
        else:
            current_trophy = row[1].strip()
            row[1] = current_trophy
            newrow = []
            for i in range(0,16):
                newrow.append('')
            newrow[0] = row[0] # Event #
            newrow[1] = row[4] # Trophy
            newrow[2] = row[2] # Event title
            newrow[3] = row[3] # Heat
            place = 1
            for i in range(5,len(row),1):
                if row[i].strip() != "":
                    (boat,time,lane) = row[i].split('\n')
                    newrow[4] = place
                    newrow[5] = boat
                    newrow[6] = time
                    place += 1
                else:
                    newrow[4] = ''
                    newrow[5] = ''
                    newrow[6] = ''

                #print([YEAR] + newrow)
                outwriter.writerow([YEAR] + newrow)


infile.close()
outfile.close()
os.unlink(CSV)
