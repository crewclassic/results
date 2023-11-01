import re
import csv

input_file = "../2019.txt"
csv_file = "2019-etl.csv"

# Initialize variables to store parsed data
parsed_data = []
current_event = None
header_names = None

# Define regular expressions for matching event number, event name, and data rows
event_pattern = re.compile(r'#(\d+) (.+)')
header_pattern = re.compile(r'Place Entry Lane Time Margin')
data_pattern = re.compile(r'(\d+) (.+?) (\d+) ([\d:.]+)(?: (\d+\.\d+))?')

# Read the data
with open(input_file, 'r') as file:
    lines = file.read().split('\n')

# Skip the first two rows of the data
lines = lines[2:]

for line in lines:
    if not line:
        continue  # Skip empty lines

    event_match = event_pattern.match(line)
    header_match = header_pattern.match(line)

    if event_match:
        # Start of a new event
        if current_event:
            parsed_data.append(current_event)  # Save the previous event
        current_event = {}
        current_event["Event #"] = event_match.group(1)
        current_event["Event"]   = event_match.group(2)
        continue

    if header_match:
        # Header row, extract column headers
        header_names = line.split()
        continue

    if current_event and header_names:
        # Data row, parse and add to the current event
        data_match = data_pattern.match(line)
        if data_match:
            data = data_match.groups()
            row = {
                "Event #": current_event["Event #"],
                "Trophy" : "",
                "Event": current_event["Event"],
                "Heat" : "",
                "Place": data[0],
                "Boat": data[1],
                "Time": data[3],
            }
            parsed_data.append(row)

# Write the parsed data to a CSV file

with open(csv_file, "w", newline="") as csvfile:
#    fieldnames = ["Event number", "Event name", "Place", "Entry", "Lane", "Time", "Margin"]
    fieldnames = ["Event #", "Trophy", "Event", "Heat", "Place", "Boat", "Time"]
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(parsed_data)

print(f"Parsed data saved to {csv_file}")

