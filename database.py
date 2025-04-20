import json
import sqlite3
import datetime
import os

def sync_json_to_sqlite(json_file='id_mapping.json', db_file='id_mapping.db'):
    if not os.path.exists(json_file):
        print(f"[!] File {json_file} tidak ditemukan.")
        return

    with open(json_file, 'r') as f:
        try:
            data = json.load(f)
        except json.JSONDecodeError:
            print(f"[!] Format JSON tidak valid.")
            return

    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS id_mapping (
            id TEXT PRIMARY KEY,
            value TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS id_mapping_log (
            id TEXT,
            old_value TEXT,
            new_value TEXT,
            action TEXT,
            timestamp TEXT
        )
    ''')

    for id_key, value in data.items():
        cursor.execute('SELECT value FROM id_mapping WHERE id = ?', (id_key,))
        row = cursor.fetchone()

        if row is None:
            cursor.execute('INSERT INTO id_mapping (id, value) VALUES (?, ?)', (id_key, value))
            cursor.execute('INSERT INTO id_mapping_log VALUES (?, ?, ?, ?, ?)',
                           (id_key, None, value, 'insert', datetime.datetime.now().isoformat()))
        elif row[0] != value:
            old_value = row[0]
            cursor.execute('UPDATE id_mapping SET value = ? WHERE id = ?', (value, id_key))
            cursor.execute('INSERT INTO id_mapping_log VALUES (?, ?, ?, ?, ?)',
                           (id_key, old_value, value, 'update', datetime.datetime.now().isoformat()))

    cursor.execute('SELECT id FROM id_mapping')
    db_ids = {row[0] for row in cursor.fetchall()}
    json_ids = set(data.keys())

    for deleted_id in db_ids - json_ids:
        cursor.execute('SELECT value FROM id_mapping WHERE id = ?', (deleted_id,))
        old_value = cursor.fetchone()[0]
        cursor.execute('DELETE FROM id_mapping WHERE id = ?', (deleted_id,))
        cursor.execute('INSERT INTO id_mapping_log VALUES (?, ?, ?, ?, ?)',
                       (deleted_id, old_value, None, 'delete', datetime.datetime.now().isoformat()))

    conn.commit()
    conn.close()
