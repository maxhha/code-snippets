"""
These programm converts ods files to json. Usage:
odf2json inputfile.ods outputfile.json sheet_number

Example: 
If table looks like this:
  A      B      C      D     E
1 Prop1  Prop2  Prop3        Prop4
2        hello  world  !!    !
3 1000   =A3*C3 6,50%  10,0  0,00%
4   
5 and    test   data

JSON is:

[{"Prop2": "hello", "Prop3": "world"},
{"Prop1": 1000.0, "Prop2": 65.0, "Prop3": 0.065}]

"""

import ezodf
import json
from pathlib import Path
from sys import argv

assert len(argv) == 4
assert Path(argv[1]).is_file(), "Read file does not exsist" 

input_file_name = argv[1]
output_file_name = argv[2]
sheet_number = int(argv[3])

doc = ezodf.opendoc(input_file_name)

sheet = doc.sheets[0]
headers = []

for cell in sheet.row(0):
	if cell.value != None:
		headers.append(cell.value)
	else:
		break

data = []
for i in range(1, sheet.nrows()):
	row = sheet.row(i)
	record = {}
	for col, cell in enumerate(row):
		if col >= len(headers):
			break
		if cell.value != None:
			record[headers[col]] = cell.value
	if len(record) > 0:
		data.append(record)
	else:
		break

with open(output_file_name, "w") as f:
	json.dump(data, f)
