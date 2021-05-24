# Write out dates for well scheduling in GAP

# David Taylor, 06th May 2013

# Set up Win32 coms

import win32com, win32com.client, csv

from win32com.client import Dispatch
	
GAP = Dispatch("PX32.Openserver.1")

# Set abandonment rate

abrat = 100

# Get well count for going through all wells

wells = GAP.GetValue("GAP.MOD[0].WELL.COUNT")
well_count = int(wells) # Convert from string to integer

# For each well name, loop through the list and write the start date and well status		

for x in range(well_count):
		
	# Get well names

        get_well_string = 'GAP.MOD[{PROD}].WELL[%s].Label' % x
        well = GAP.GetValue(get_well_string)

        setabrat='GAP.MOD[{PROD}].WELL[{%s}].IPR[0].ABMinQoil' % (well)
        GAP.SetValue(setabrat, abrat)

		
	



