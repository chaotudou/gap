# Get network solver oil and injected gas rates from GAP
# David Taylor, 26th October 2011

# Set up Win32 coms

import win32com, win32com.client

from win32com.client import Dispatch
	
GAP = Dispatch("PX32.Openserver.1")

# Make output file for data - Change directory to wherever file should be saved
	
outfile = open ("H:\GAP_prediction.txt", "w")

# Set unit warning

unitwarn = "**** CHECK THAT UNITS IN GAP CORRESPOND TO UNITS BELOW *****"
outfile.write(unitwarn)
outfile.write("\n")

# Get run name

run_name = GAP.GetValue("GAP.MOD[0].FILENAME")
run_name_list = run_name.split("\\")
run_name = run_name_list.pop()

outfile.write("\n")	
outfile.write("File name:")
outfile.write(run_name)
outfile.write("\n")

# Write column headers

header = "\nWell\tOil Rate(bbl/day)\tGas Rate(MMscf/day)\tWatercut (%)\n\n"
outfile.write(header)

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

# Get oil rate

        qo_rate_string = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{31/03/2013}].OILRATE' % (well)
        qo_rate = GAP.GetValue(qo_rate_string)

# Get gas rate

        gl_rate_string = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{31/03/2013}].GASRATE' % (well)
        gl_rate = GAP.GetValue(gl_rate_string)

# Get watercut

        wct_string = 'GAP.MOD[{PROD}].WELL[{%s}].PREDRES[{30/03/2016}].WCT' % (well)
        wct = GAP.GetValue(wct_string)

# Remove | at the end of qo_cum_end string (not sure why this appears)

        if qo_rate.endswith('|'):
                qo_cum_end = qo_cum_end[:-1]

        outfile.write(well)
        outfile.write('\t')
        outfile.write(qo_rate)
        outfile.write('\t')
        outfile.write(gl_rate)
        outfile.write('\t')
        outfile.write(wct)
        outfile.write('\n')
        
