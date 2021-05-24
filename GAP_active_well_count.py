# Write out active well counts in a specific GAP project

# David Taylor, 17th July 2013

# Set up Win32 coms

import win32com, win32com.client, csv

from win32com.client import Dispatch

GAP = Dispatch("PX32.Openserver.1")

# Get well count to loop over all wells

wells = GAP.GetValue("GAP.MOD[0].WELL.COUNT")
well_count = int(wells) # Convert from string to integer

# Get date count

datecount = 212

# Get dates from file

filein=open("H:\dates.txt","rb")
outfile = open("H:\wellcount.txt", "w")

# Define active well count

activecount = 0

for lines in filein:

        for y in range(well_count):
        
        # Get well name for current loop
        
            get_well_string = 'GAP.MOD[{PROD}].WELL[%s].Label' % y
            well = GAP.GetValue(get_well_string)
            date = lines
            orat = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{%s}].OILRATE' % (well, date)
            get_orat = GAP.GetValue(orat)

            try:    
                    int_orat = int(float(get_orat))
            except:
                    int_orat = 0
                    
        # Check if well is active

            if int_orat > 0:
                    activecount = activecount + 1
        
        outfile.write(str(activecount))
        outfile.write('\n')
        activecount = 0
    
        
        


		
	



