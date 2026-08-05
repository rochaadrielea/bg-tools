import sqlite3, pandas as pd
con = sqlite3.connect("quality.db")
df = pd.read_sql_query(open("db_query.sql").read(), con)
print(df)
con.close()