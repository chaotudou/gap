# Get cumulative oil production for specific wells from GAP
# David Taylor, 11th August 2011

# Set up Win32 coms

import win32com, win32com.client

from win32com.client import Dispatch
	
GAP = Dispatch("PX32.Openserver.1")

# Make output file for data - Change directory to wherever file should be saved
	
outfile = open ("H:/GAP_cum_oil.txt", "w")

# Get run name

run_name = GAP.GetValue("GAP.MOD[0].FILENAME")
run_name_list = run_name.split("\\")
run_name = run_name_list.pop()
	
outfile.write(run_name)
outfile.write("\n")

# Write column headers

header = "\nWell\tCum Oil(MMbbl)\n\n"
outfile.write(header)

# Get prediction end date

pred_end = GAP.GetValue("GAP.PREDINFO.END.DATESTR")
y = pred_end

# Get well list

wells = GAP.GetValue("GAP.MOD[0].WELL.COUNT")
well_count = int(wells) # Convert from string to integer

# Loop over wells

for x in range(well_count):

## Filter out masked wells - unquote if necessary
##
##    well_status_string = 'GAP.MOD[{PROD}].WELL[%s].ISMASKED' % x
##    well_status = GAP.GetValue(well_status_string)
##
##    if well_status == 0:

# Get well names

        get_well_string = 'GAP.MOD[{PROD}].WELL[%s].Label' % x
        well = GAP.GetValue(get_well_string)

# Get cumulative oil

        qo_cum_end_string = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{%s}].CUMOIL' % (well, y)
        qo_cum_end = GAP.GetValue(qo_cum_end_string)

# Remove | at the end of qo_cum_end string (not sure why this appears)

        if qo_cum_end.endswith('|'):
                qo_cum_end = qo_cum_end[:-1]

        outfile.write(well)
        outfile.write('\t')
        outfile.write(qo_cum_end)
        outfile.write('\n')
        
