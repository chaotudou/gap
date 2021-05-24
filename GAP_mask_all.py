# Mask all wells in a GAP project
# David Taylor, 14th May 2013

# Set up Win32 coms

import win32com, win32com.client

from win32com.client import Dispatch
	
GAP = Dispatch("PX32.Openserver.1")

# Get well list

wells = GAP.GetValue("GAP.MOD[0].WELL.COUNT")
well_count = int(wells) # Convert from string to integer

# Loop over wells

for x in range(well_count):

# Get well names

        get_well_string = 'GAP.MOD[{PROD}].WELL[%s].Label' % x
        well = GAP.GetValue(get_well_string)

        mask_well_string = 'GAP.MOD[{PROD}].WELL[{%s}].ISMASKED()' % well
        mask_well = GAP.SetValue(mask_well_string,'0')   
