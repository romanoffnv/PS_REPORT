# 1_acc.py

1. Getting all data from the column 1 (mols **82** and items **2111**) and filtering mols out of it
2. Correlating the number of mols to match the number of items (2029-2029)
3. Post the data acquired into the db as accountance_1

# 2_acc.py

1. Get accountance_1 db cols as lists (L_mols and L_items)
2. Build dataframe
3. Populate list of keywords for df filtration (to have vehicles only)
4. Filter df by keywords
5. Destructure df into lists
6. Derive the list of untouchable units to post it into db later
7. Slice items starting from the keyword's index to the end of the sentence
8. Remove paranthesis and content in it
9. Remove crap like 'г/н' etc
10. Fishing out plates by regex from long sentences
11. Remove pointless sentences
12. Remove regions
13. Make some crutches
14. Bring plates to 111abc format
15. Push lists to DB as accountance_2

# 3_acc.py

1. Connect to omnicomm.db get Final_DB as L_om_index_plates
2. Connect to accountance.db get accountance_2 (mols, items) as L_mols and L_acc_index_plates
3. Get lists L_mols and L_acc_index_plates into dict
4. Match if Omnicomm items are in accountance
5. Push L_mols_matched into final_DB as Responsible
6. Match if accountance items are not in Omnicomm
7. Get df of unmatched items
8. Filter df for unmatched items
9. Destructuring data into lists
10. Push the lists to db
