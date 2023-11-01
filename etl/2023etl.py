#!/usr/bin/python3

import subprocess
import csv
import os, sys
from chardet import detect

YEAR="2023"
WB="../{}/2023-results.xlsx".format(YEAR)
CSV="{}.csv".format(YEAR)
ETL="{}-etl.csv".format(YEAR)

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

# 2023 format:
# header on line 8
#   bowNumber, eventCode, eventName, entryName, avgAge, strokeName, orgName, officialTime, resultStatus, eventPercent, top5Percent
# For each event, the rows are in order from 1st to last


# Output:
# Year, Event #, Trophy, Event, Heat, Place, Organization, Time


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
            if row[0] == "bowNumber":
                state = 1
        elif state == 1:
            event_number = row[1]
            if last_event_number != event_number:
                place = 1
                last_event_number = event_number
            event_name   = row[2]
            boat_name    = row[3]
            stroke_name  = row[5]
            time         = row[7]
            
            newrow = []
            for i in range(0,16):
                newrow.append('')

            newrow[0] = event_number
            newrow[1] = '' # 2022 data has trophy names in the event name
            newrow[2] = event_name
            newrow[3] = '' # 2022 heat is in the name
            newrow[4] = place
            place += 1
            newrow[5] = boat_name
            newrow[6] = time
            newrow[8] = stroke_name

            #print([YEAR] + newrow)
            outwriter.writerow([YEAR] + newrow)


infile.close()
outfile.close()
os.unlink(CSV)
