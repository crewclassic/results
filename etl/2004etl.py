#!/usr/bin/python3

import subprocess
import csv
import os, sys


YEAR="2004"
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
event = ""
current_event = ""

with open(CSV, 'r') as infile:
    inreader = csv.reader(infile, dialect='excel')
    for row in inreader:
        if(rowcount == 0):
            row = ["Year", "Event #", "Trophy", "Event", "Heat", "Place", "Boat", "Time", "Cox", "Rower1", "Rower2", "Rower3", "Rower4", "Rower5", "Rower6", "Rower7", "Rower8"]
            outwriter.writerow(row)
        else:
            newrow = []
            for i in range(0,16):
                newrow.append('')

            if(event == "" and row[3] != ""):
                event = row[3]
                current_event = event
            elif (event != "" and row[3] == ""):
                current_event = event
            else:
                current_event = row[3]
                event = current_event
            newrow[0] = '' # no event #'s
            newrow[1] = row[2]  # trophy
            newrow[3] = current_event  # heat
            newrow[4] = row[1]   # place
            newrow[5] = row[4]
            newrow[6] = row[5]
            for i in range(7,16):
                newrow[i] = row[i-1]
            
            outwriter.writerow([YEAR] + newrow)

        rowcount += 1

infile.close()
outfile.close()
os.unlink(CSV)
