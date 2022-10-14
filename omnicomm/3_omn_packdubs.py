
import sqlite3
from pprint import pprint
from collections import defaultdict
import pandas as pd


db = sqlite3.connect('omnicomm.db')
cursor = db.cursor()

def main():
    query = db.execute("SELECT * FROM total_vehicles")
    cols = [column[0] for column in query.description]
    data = pd.DataFrame.from_records(data = query.fetchall(), columns = cols)
    
    L_total_dep = data['Department'].tolist()
    L_total_veh = data['Vehicle'].tolist()
    L_total_veh = [x.strip() for x in L_total_veh]
    

##Assigning duplicated groups to the plate numbers
    dc = defaultdict(list)

    for i in range(len(L_total_veh)):
        item = L_total_veh[i]
        dc[item].append(L_total_dep[i])

    dc = dict(zip(dc.keys(), map(set, dc.values())))
    
    L_total_veh, L_total_dep  = zip(*dc.items())
    L_total_dep = [', '.join(x) for x in L_total_dep]
    
    data = pd.DataFrame(zip(L_total_dep, L_total_veh), columns =['Department', 'Vehicle'])
    
    
    # Updating table of total vehicles with red column
    
    cursor.execute("DROP TABLE IF EXISTS total_vehicles_nobubs;")
    cursor.execute("""
                    CREATE TABLE IF NOT EXISTS total_vehicles_nobubs(
                    Department text,
                    Vehicle text
    
                                )
                    """)
    cursor.executemany("INSERT INTO total_vehicles_nobubs VALUES (?, ?)", zip(data.loc[:, 'Department'], data.loc[:, 'Vehicle']))
    print(len(cursor.execute("SELECT * FROM total_vehicles_nobubs").fetchall()))
    
    # refreshing database
    db.commit()
    # closing database
    db.close()
    
    print("3_omn_packdubs is complete")
if __name__ == '__main__':
    main()
