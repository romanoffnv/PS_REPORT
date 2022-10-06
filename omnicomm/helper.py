import sqlite3
db = sqlite3.connect('omnicomm.db')
cursor = db.cursor()
cursor.execute("DROP TABLE IF EXISTS final_DB;")