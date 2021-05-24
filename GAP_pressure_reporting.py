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

datecount = 1

# Open a file to store the well counts in

outfile = open("H:\wellpressures.txt", "w")

# Define active well count object

for y in range(well_count):
    get_well_string = 'GAP.MOD[{PROD}].WELL[%s].Label' % y
    well = GAP.GetValue(get_well_string)
    outfile.write(well)
    outfile.write('\t')
    outfile.write('\t')
    outfile.write('\t')

outfile.write('\n')

for y in range(well_count):
    outfile.write("THP (psig)")
    outfile.write('\t')
    outfile.write("BHP (psig)")
    outfile.write('\t')
    outfile.write("Reservoir Pressure (psig)")
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
        day = day
        month = short_list[1]
        year = short_list[0]
        if year > 2013 and day > 28:
            day = str(28)
        else:
            day = str(day)
        month = str(short_list[1])
        year = str(short_list[0])
        date = day+'/'+month+'/'+year # There is probably a better way to do this...
            
        # Get well name for current loop
        
        get_well_string = 'GAP.MOD[{PROD}].WELL[%s].Label' % y
        well = GAP.GetValue(get_well_string)

        if well == "NSL_III/IV_I":
            pass
        else:
            continue

        resp = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{%s}].RESPRES' % (well, date)
        try:
            get_resp = float(GAP.GetValue(resp))
        except:
            get_resp = 0
        thp = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{%s}].FWHP' % (well, date)
        try:
            get_thp = float(GAP.GetValue(thp))
        except:
            get_thp = 0
        bhp = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{%s}].FBHP' % (well, date)
        try:
            get_bhp = float(GAP.GetValue(bhp))
        except:
            get_bhp = 0
       
        if get_resp > 10000:
            get_resp = ""
          
        if get_thp > 10000:
            get_thp = ""
        
        if get_bhp > 10000:
            get_bhp = ""
        
        outfile.write(get_thp)
        outfile.write('\t')
        outfile.write(get_bhp)
        outfile.write('\t')
        outfile.write(get_resp)
        outfile.write('\t')
    outfile.write('\n')
        
    
        
        


		
	



