import os
import sqlite3
import psycopg2
from psycopg2.extras import RealDictCursor

SQLITE_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'database.db')
DATABASE_URL = os.environ.get('DATABASE_URL', '').strip()

TABLES = {
    'users': ['id','name','email','password','role','created_at'],
    'students': ['id','name','dni','code','entry_date','active','created_at','updated_at'],
    'teachers': ['id','name','dni','area','active','created_at','updated_at'],
    'searches': ['id','user_id','query','result','student_name','created_at'],
    'access_logs': ['id','user_id','student_id','query','result','student_name','student_dni','student_code','note','created_at'],
    'teacher_logs': ['id','user_id','teacher_id','query','result','teacher_name','teacher_dni','teacher_area','note','created_at'],
    'visitors': ['id','user_id','full_name','dni','destination_area','reason','attended','attended_at','created_at'],
}

if not DATABASE_URL:
    raise SystemExit('ERROR: primero coloca la variable DATABASE_URL de PostgreSQL.')
if not os.path.exists(SQLITE_DB):
    raise SystemExit(f'ERROR: no encontré {SQLITE_DB}')

sq = sqlite3.connect(SQLITE_DB)
sq.row_factory = sqlite3.Row
pg = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
cur = pg.cursor()

import app
app.init_db()

for table, cols in TABLES.items():
    rows = sq.execute(f'SELECT {",".join(cols)} FROM {table}').fetchall()
    if not rows:
        print(f'{table}: 0 filas')
        continue
    placeholders = ','.join(['%s'] * len(cols))
    updates = ','.join([f'{c}=EXCLUDED.{c}' for c in cols if c != 'id'])
    sql = f'''INSERT INTO {table} ({','.join(cols)}) VALUES ({placeholders})
              ON CONFLICT (id) DO UPDATE SET {updates}'''
    for r in rows:
        cur.execute(sql, [r[c] for c in cols])
    cur.execute(f"SELECT setval(pg_get_serial_sequence('{table}', 'id'), COALESCE((SELECT MAX(id) FROM {table}), 1), true)")
    print(f'{table}: {len(rows)} filas migradas')

pg.commit()
cur.close(); pg.close(); sq.close()
print('Migración terminada. Tus datos de database.db ya están en PostgreSQL.')
