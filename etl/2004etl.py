#!/usr/bin/python3

import subprocess
import csv
import os, sys


WB="../2004.xls"
CSV="2004.csv"
ETL="2004-etl.csv"

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
            outwriter.writerow(row)
        else:
            if(event == "" and row[3] != ""):
                event = row[3]
                current_event = event
            elif (event != "" and row[3] == ""):
                current_event = event
            else:
                current_event = row[3]
                event = current_event
            row[3] = current_event
            outwriter.writerow(row)

        rowcount += 1

infile.close()
outfile.close()
os.unlink(CSV)
