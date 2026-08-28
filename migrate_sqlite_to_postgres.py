import os
import sqlite3
import psycopg2
from psycopg2.extras import RealDictCursor

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SQLITE_DB = os.path.join(APP_DIR, 'database.db')
DATABASE_URL = os.environ.get('DATABASE_URL', '').strip()
TABLES = ['users', 'students', 'teachers', 'searches', 'access_logs', 'teacher_logs', 'visitors', 'audit_logs', 'registration_requests']

if not DATABASE_URL:
    raise SystemExit('ERROR: configura DATABASE_URL antes de ejecutar la migración.')
if not os.path.exists(SQLITE_DB):
    raise SystemExit(f'ERROR: no se encontró {SQLITE_DB}')

# app.init_db() crea/migra la estructura multisede en PostgreSQL.
import app
app.init_db()

sq = sqlite3.connect(SQLITE_DB)
sq.row_factory = sqlite3.Row
pg = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
cur = pg.cursor()

for table in TABLES:
    sqlite_cols = [r['name'] for r in sq.execute(f'PRAGMA table_info({table})').fetchall()]
    if not sqlite_cols:
        print(f'{table}: tabla no presente en SQLite')
        continue
    cur.execute('''SELECT column_name FROM information_schema.columns
                   WHERE table_schema='public' AND table_name=%s ORDER BY ordinal_position''', (table,))
    postgres_cols = [r['column_name'] for r in cur.fetchall()]
    cols = [col for col in sqlite_cols if col in postgres_cols]
    rows = sq.execute(f'SELECT {",".join(cols)} FROM {table}').fetchall()
    if not rows:
        print(f'{table}: 0 filas')
        continue
    placeholders = ','.join(['%s'] * len(cols))
    updates = ','.join([f'{col}=EXCLUDED.{col}' for col in cols if col != 'id'])
    conflict = f'ON CONFLICT (id) DO UPDATE SET {updates}' if 'id' in cols and updates else ''
    sql = f'INSERT INTO {table} ({",".join(cols)}) VALUES ({placeholders}) {conflict}'
    for row in rows:
        cur.execute(sql, [row[col] for col in cols])
    if 'id' in cols:
        cur.execute(f"SELECT setval(pg_get_serial_sequence('{table}', 'id'), COALESCE((SELECT MAX(id) FROM {table}), 1), true)")
    print(f'{table}: {len(rows)} filas migradas')

pg.commit()
cur.close(); pg.close(); sq.close()
print('Migración multisede terminada correctamente.')
