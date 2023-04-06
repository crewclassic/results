#!/usr/bin/python3

import subprocess
import csv
import os, sys
from chardet import detect

YEAR="2018"
WB="../{}.xlsx".format(YEAR)
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

# 2018 format:
# Details, Event ID, Event Name, Race ID, Race No, Date, Time, Entries, Distance, Status, Progression Rule

# Output:
# Year, Event #, Trophy, Event, Heat, Place, Boat, Time, Cox, Rower1, Rower2, Rower3, Rower4, Rower5, Rower6, Rower7, Rower8


with open(CSV, 'r', encoding = get_encoding_type(CSV), errors='ignore') as infile:
    inreader = csv.reader(infile, dialect='excel')

    row = ["Year", "Event #", "Trophy", "Event", "Heat", "Place", "Boat", "Time", "Cox", "Rower1", "Rower2", "Rower3", "Rower4", "Rower5", "Rower6", "Rower7", "Rower8"]
    outwriter.writerow(row)

    state = 0
    adj_time_index = 0
    for row in inreader:
        rowcount += 1

        if state == 0 and rowcount != 368:
            # Locate an event row
            if len(row) >= 321:
                if row[321] == "Official":
                    state = 1
                    newrow = []
                    for i in range(0,16):
                        newrow.append('')

                    newrow[0] = row[2] # event #
                    newrow[1] = ''     # 2018 trophy data embedded in the event name
                    newrow[2] = row[4] # Event title
                    heat = row[61]     # Heat
                    newrow[3] = heat.replace("Heat ","")
                    
                    num_entries = int(row[240])
        elif state == 1:
            if len(row) >= 1:
                if row[1] == "Detail":
                    state = 2
                    adj_time_index = row.index("Adj Time")
                    description_index = row.index("Description")
        elif state == 2:
            if len(row) >= 470:
                place = row[2]
                boat = row[5] + " " + row[46]
                time = row[adj_time_index]
            
                newrow[4] = place.strip()
                newrow[5] = boat.strip()
                newrow[6] = time.strip()
                newrow[8] = row[description_index].strip()
            
                #print([YEAR] + newrow)
                outwriter.writerow([YEAR] + newrow)
                
                newrow[4] = ''
                newrow[5] = ''
                newrow[6] = ''
                num_entries -= 1
                if num_entries == 0:
                    state = 0
            elif len(row) == 0:
                state = 0
                
infile.close()
outfile.close()
os.unlink(CSV)
