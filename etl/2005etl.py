#!/usr/bin/python3

import subprocess
import csv
import os, sys


YEAR="2005"
WB="../{}.xls".format(YEAR)
CSV="{}.csv".format(YEAR)
ETL="{}-etl.csv".format(YEAR)

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
# Event #, Trophy, Title, Heat, Start, 1st Place Organization, 1st Place time... 7th Place Organization, 7th Place time
# Output:
# Year, Event #, Trophy, Event, Heat, Place, Boat, Time, Cox, Rower1, Rower2, Rower3, Rower4, Rower5, Rower6, Rower7, Rower8

with open(CSV, 'r') as infile:
    inreader = csv.reader(infile, dialect='excel')
    for row in inreader:
        row[1] = row[1].strip() # Trophy name
        row[2] = row[2].strip() # Event Title

        if(rowcount == 2):
            row = ["Year", "Event #", "Trophy", "Event", "Heat", "Place", "Boat", "Time", "Cox", "Rower1", "Rower2", "Rower3", "Rower4", "Rower5", "Rower6", "Rower7", "Rower8"]
            outwriter.writerow(row)
        elif(rowcount >= 56 and rowcount <= 59 or rowcount < 2):
            next
        else:
            if "final" in row[3].lower() or "gfinal" in row[3].lower():
                current_trophy = row[1].strip()
            else:
                current_trophy = ""
            row[1] = current_trophy
            newrow = row[0:4]    # Event #, Trophy, Event, Heat

            for i in range(1,13):
                newrow.append('')
            place = 1
            for i in range(5,len(row),2):
                newrow[4] = place
                newrow[5] = row[i]           # boat
                newrow[6] = row[i+1].strip()  # time
                place += 1
                #print([YEAR] + newrow)
                outwriter.writerow([YEAR] + newrow)

        rowcount += 1

infile.close()
outfile.close()
os.unlink(CSV)
