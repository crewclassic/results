#!/usr/bin/python3

import subprocess
import csv
import os, sys
from chardet import detect

WB="../2008.xls"
CSV="2008.csv"
ETL="2008-etl.csv"

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

# Output:
# Event #, Trophy, Event, Heat, Place, Organization, Time


with open(CSV, 'r', encoding = get_encoding_type(CSV), errors='ignore') as infile:
    inreader = csv.reader(infile, dialect='excel')

    for row in inreader:
        rowcount += 1

        if(rowcount >= 60 and rowcount <= 62 or rowcount < 3 or rowcount >= 99):
            continue

        if(rowcount == 3):
            row = ["Event #", "Trophy", "Event", "Heat", "Place", "Boat", "Time", "Cox", "Rower1", "Rower2", "Rower3", "Rower4", "Rower5", "Rower6", "Rower7", "Rower8"]
            outwriter.writerow(row)
        else:
            current_trophy = row[1].strip()
            row[1] = current_trophy
            newrow = []
            for i in range(0,18):
                newrow.append('')
            newrow[0] = row[0] # Event #
            newrow[1] = row[4] # Trophy
            newrow[2] = row[2] # Event title
            newrow[3] = row[3] # Heat
            place = 1
            for i in range(5,len(row),1):
                if row[i].strip() != "":
                    try:
                        (boat,time,lane) = row[i].split('\n')
                    except:
                        print("can't split >{}<".format(row[i]))

                    newrow[4] = place
                    newrow[5] = boat
                    newrow[6] = time
                    place += 1
                else:
                    newrow[4] = ''
                    newrow[5] = ''
                    newrow[6] = ''

                print(newrow)
                outwriter.writerow(newrow)


infile.close()
outfile.close()
os.unlink(CSV)