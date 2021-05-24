# Write out active well counts in a specific GAP project

# David Taylor, 17th July 2013

# Set up Win32 coms

import win32com, win32com.client, csv, xlrd

from win32com.client import Dispatch

GAP = Dispatch("PX32.Openserver.1")

# Get well count to loop over all wells

wells = GAP.GetValue("GAP.MOD[0].WELL.COUNT")
well_count = int(wells) # Convert from string to integer

# Get date count (ctrl+right click on the last date in GAP prediction results to find this number)

datecount = 282

# Open a file to store the well counts in

outfile = open("H:\wellproduction.txt", "w")

# Define active well count object

for y in range(well_count):
    get_well_string = 'GAP.MOD[{PROD}].WELL[%s].Label' % y
    well = GAP.GetValue(get_well_string)
    outfile.write(well)
    outfile.write('\t')
    outfile.write('\t')
    outfile.write('\t')
    outfile.write('\t')

outfile.write('\n')

for y in range(well_count):
    outfile.write("Oil Rate (bbl/d)")
    outfile.write('\t')
    outfile.write("Gas Rate (mmscf/d)")
    outfile.write('\t')
    outfile.write("Water Rate (bbl/d)")
    outfile.write('\t')
    outfile.write("Gas Lift Rate (mmscf/d")
    outfile.write('\t')
   
outfile.write('\n')
		
for x in range(datecount):

    for y in range(well_count):

        # Get timestep

        get_date = 'GAP.MOD[{PROD}].WELL[{N11}].PREDRES.DATES[%s]' % x
        date = GAP.GetValue(get_date) # This returns and Excel format date and so must be converted back into a string of the format dd/mm/yyyy
        date = float(date)
        convert_date = xlrd.xldate.xldate_as_tuple(date,0) # Converts an Excel date format to Python date format
        date_list = list(convert_date)
        short_list = [date_list[i] for i in (0,1,2)] # Removes any time measurement smaller than a day
        day = short_list[2] + 1 # For some reason the Excel date was coming back one day short so this corrects that
        month = short_list[1]
        year = short_list[0]
        if day > 31:
                day = 01
                month = short_list[1] + 1
        if day > 30 and month is 9:
                day = 01
                month = 10
        if day > 28 and month is 2:
                day = 01
                month = 03
        if month is 13:
                day = 01
                month = 01
                year = short_list[0] + 1
        if month is 6 and day > 30:
                day = 01
                month = 07
                #if year > 2013 and day > 28 and month > 01:
                #        day = 28
        day = str(day)
        month = str(month)
        year = str(year)
        date = day+'/'+month+'/'+year # There is probably a better way to do this...
            
        # Get well name for current loop
        
        get_well_string = 'GAP.MOD[{PROD}].WELL[%s].Label' % y
        well = GAP.GetValue(get_well_string)

        orat = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{%s}].OILRATE' % (well, date)
        get_orat = GAP.GetValue(orat)
        
        grat = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{%s}].GASRATE' % (well, date)
        get_grat = GAP.GetValue(grat)
        
        wrat = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{%s}].WATRATE' % (well, date)
        get_wrat = GAP.GetValue(wrat)

        glinj = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{%s}].GASLIFTINJ' % (well, date)
        get_glinj = GAP.GetValue(glinj)
                      
        outfile.write(get_orat)
        outfile.write('\t')
        outfile.write(get_grat)
        outfile.write('\t')
        outfile.write(get_wrat)
        outfile.write('\t')
        outfile.write(get_glinj)
        outfile.write('\t')
    outfile.write('\n')
        
    
        
        


		
	



