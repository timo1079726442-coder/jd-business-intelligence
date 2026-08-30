import sqlite3
db=sqlite3.connect('data/jd_report.db')
for (name,) in db.execute("select name from sqlite_master where type='table' order by name"):
    cols=db.execute(f'pragma table_info("{name}")').fetchall()
    print(name, ' | '.join(f'{c[1]}:{c[2]}' for c in cols))
