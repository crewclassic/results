#!/usr/bin/python3

import subprocess
import csv
import os, sys
from chardet import detect

YEAR="2014"
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

# Output:
# Year, Event #, Trophy, Event, Heat, Place, Organization, Time


with open(CSV, 'r', encoding = get_encoding_type(CSV), errors='ignore') as infile:
    inreader = csv.reader(infile, dialect='excel')

    for row in inreader:
        rowcount += 1

        if(rowcount < 3 or rowcount >= 67 and rowcount <= 69 or rowcount > 116):
            continue

        if(rowcount == 3):
            row = ["Year", "Event #", "Trophy", "Event", "Heat", "Place", "Boat", "Time", "Cox", "Rower1", "Rower2", "Rower3", "Rower4", "Rower5", "Rower6", "Rower7", "Rower8"]
            outwriter.writerow(row)
        else:
            if event == "" and row[2].strip() != "":
                event = row[2].strip()
            elif event != "" and row[2].strip() == "":
                pass
            elif event != row[2].strip():
                event = row[2].strip()

            if "final" in row[3].lower() or "gfinal" in row[3].lower():
                current_trophy = row[4].strip()
            else:
                current_trophy = ""

            row[1] = current_trophy
            newrow = []
            for i in range(0,16):
                newrow.append('')
            newrow[0] = row[0] # Event #
            newrow[1] = current_trophy # Trophy
            newrow[2] = event # Event title
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

                #print([YEAR] + newrow)
                outwriter.writerow([YEAR] + newrow)


infile.close()
outfile.close()
os.unlink(CSV)
