accountance_1

accountance_1

# 1_acc.py

1. This component opens up xls file of accountance dept and forms up rough database of all items there are, outputting them into accountance_1 db
2. Posts DB shy for "Гамма плотномер" containing items (because they interfere with the trailer plates) as accountance_2

# 2_acc.py

1. This component filters out mols, units, plates from accountance_2 into df2 and merges it to df1, taken from accountance_1, then posts merged df to DB as accountance_3

# accountance_1, 2, 3

accountance_1.db is rough data base taken as is from xls file

accountance_2.db is database cleaned off "Гамма плотномер" containing items"Гамма плотномер" containing items

accountance_3.db is ultimate data base, merged of dfs from accountance_1 and accountance_2
