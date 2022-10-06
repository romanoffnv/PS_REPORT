# Get table from acc_unmached as lists
# Put the list to df
# Find properties, common for the units to be in Omnicomm (e.g. 'Vehicle', 'Tractor' etc.) and put them into the list as keywords
# Sort df by the keywords (leaving only vehicles to be in Omn)
# Get columns: Vehicle_om, Plate_index_om from omnicomm.db
# Partial match Plate_index_acc to Plate_index_om, if match is 83,33% (1 mismatch) get f"string" as descrepancy and push into list to length of Om db
# Push discrepancy list into omnicomm.db

# Rearrange names in omnicomm.db (eg. Vehicle - Vehicle_om)