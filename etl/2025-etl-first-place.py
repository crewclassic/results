#!/usr/bin/env python3

import requests
import re
import csv

# Constants
regattaId = 1551  # 2025 SDCC
#rmAuth = ("", "")
racesurl   = "https://api.regattamaster.com/dataservice.svc/vw_races"
resultsurl = "https://api.regattamaster.com/dataservice.svc/vw_tf_results"
crewlisturl = "https://api.regattamaster.com/dataservice.svc/vw_crewlist"
rmHeaders = {"RegattaId": f"{regattaId}"}
seats = ["Cox", "Bow", "2", "3", "4", "5", "6", "7", "Stroke", "Coach"]

def fetch_race_details_bulk(race_details_url, params=None, headers=rmHeaders, auth=rmAuth):
    """
    Fetch race details from the given URL using HTTP GET request.
    
    Args:
        race_details_url (str): The API endpoint URL.
        params (dict, optional): Query parameters for the request.
        headers (dict, optional): Headers for the request.
        auth (tuple, optional): Authentication credentials.

    Returns:
        dict: Parsed JSON response if successful, otherwise None.
    """
    try:
        response = requests.get(race_details_url, params=params, headers=headers, auth=auth)
        response.raise_for_status()
        return response.json()

    except requests.exceptions.HTTPError as e:
        print(f"HTTP error occurred: {e.response.status_code} - {e.response.text}")
    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {str(e)}")

    return None

def fetch_event_details_line_by_line(race_details_url, params=None, headers=rmHeaders, auth=rmAuth):
    data = fetch_race_details_bulk(race_details_url, params, headers, auth)
    if data:
        for line in data['d']:
            yield line

def fetch_race_entries_line_by_line(race_details_url, params=None, headers=rmHeaders, auth=rmAuth):
    data = fetch_race_details_bulk(race_details_url, params, headers, auth)
    if data:
        for line in data['d']:
            yield line

# Main
if __name__ == "__main__":

    file = open("2025-winners.csv", mode="w", newline="")
    output = csv.writer(file)
    
    crewSelect = {
        "$filter": "RMRSID eq {}".format(regattaId),
        "$format": "json"
    }
    crewData = fetch_race_details_bulk(crewlisturl, crewSelect)
    crewIndex = {}
    for i, person in enumerate(crewData['d']):
        entryPID = person['Pid']
        regClubId = person['regClubId']
        SeatLabel = person['SeatLabel']
        if entryPID not in crewIndex:
            crewIndex[entryPID] = {}
        if regClubId not in crewIndex[entryPID]:
            crewIndex[entryPID][regClubId] = {}
        if SeatLabel not in crewIndex[entryPID][regClubId]:
            crewIndex[entryPID][regClubId][SeatLabel] = person['displayName']

    entrySelect = {
        "$filter": "RMRSID eq {}".format(regattaId),
        "$orderby": "laneNo desc",
        "$format": "json"
    }

    for entry in fetch_race_entries_line_by_line(resultsurl, entrySelect):
        match = re.search(r"^Final(?: A)?$", entry['raceLabel'])
        if match:
            if entry['finishPlace'] == 1:
                crew = []
                for seat in seats:
                    EntryPID = entry['EntryPID']
                    regClubId = entry['regClubId']
                    if seat in crewIndex.get(EntryPID, {}).get(regClubId, {}):
#                    print(f"{seat}: {crewIndex[EntryPID][regClubId][seat]}")
                        crew.append(str(crewIndex[EntryPID][regClubId][seat]))
                    else:
                        crew.append("") # must have something in the column
                try:
                    minutes, seconds = entry['adjTime'].split(":")
                    total_seconds = int(minutes) * 60 + float(seconds)
                except:
                    total_seconds = 0

                output.writerow(["2025",entry['eventId'],entry['trophy'],entry['eventDescription'],entry['raceLabel'],entry['finishPlace'],entry['ClubName'],entry['adjTime']] + crew + [total_seconds])
