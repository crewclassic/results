#!/usr/bin/python3

import subprocess
import csv
import os, sys
from chardet import detect

WB="../2006.xls"
CSV="2006.csv"
ETL="2006-etl.csv"

def unicode_csv_reader(utf8_data, dialect=csv.excel, **kwargs):
    csv_reader = csv.reader(utf8_data, dialect=dialect, **kwargs)
    for row in csv_reader:
        yield [unicode(cell, 'ISO-8859-1') for cell in row]

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
# Event #, Trophy, Title, Heat, Start, 1st Place Organization, 1st Place time... 7th Place Organization, 7th Place time
# Output:
# Event #, Trophy, Event, Heat, Place, Organization, Time



with open(CSV, 'r', encoding = get_encoding_type(CSV), errors='ignore') as infile:
    inreader = csv.reader(infile, dialect='excel')

    for row in inreader:
        row[1] = row[1].strip() # Trophy name
        row[2] = row[2].strip() # Event Title

        if(rowcount == 2):
            row = ["Event #", "Trophy", "Event", "Heat", "Place", "Boat", "Time", "Cox", "Rower1", "Rower2", "Rower3", "Rower4", "Rower5", "Rower6", "Rower7", "Rower8"]
            outwriter.writerow(row)
        elif(rowcount >= 60 and rowcount <= 61 or rowcount < 2):
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
                print(newrow)
                outwriter.writerow(newrow)

        rowcount += 1

infile.close()
outfile.close()
os.unlink(CSV)
