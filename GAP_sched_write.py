# Write out dates for well scheduling in GAP

# David Taylor, 06th May 2013

# Set up Win32 coms

import win32com, win32com.client, csv

from win32com.client import Dispatch
	
GAP = Dispatch("PX32.Openserver.1")

# Read in data file containing well names and start dates
# File format should be .csv with the first column being the well name (this must match
# with the GAP well name) and the second column being the start date

filein=open("H:\start_dates.csv","rb")

# Use the csv mnodule to read the file and split it into two lists

with filein as f:
		reader=csv.reader(f)
		xs, ys = zip(*reader)

# For each well name, loop through the list and write the start date and well status		
		
for x,y in zip(xs, ys):
		startdate=y
		closedate="31/03/2013"
		setclosedate='GAP.MOD[{PROD}].WELL[{%s}].SCHEDULE[1].Time' %x
		closewell='GAP.MOD[{PROD}].WELL[{%s}].SCHEDULE[1].TYPE' %x
		setwelldate='GAP.MOD[{PROD}].WELL[{%s}].SCHEDULE[0].Time' %x
		setwellstart='GAP.MOD[{PROD}].WELL[{%s}].SCHEDULE[0].TYPE' %x
		GAP.SetValue(setwelldate, startdate)
		GAP.SetValue(setwellstart, "WELL_ON")
		GAP.SetValue(setclosedate, closedate)
		GAP.SetValue(closewell, "WELL_OFF")
		
	



