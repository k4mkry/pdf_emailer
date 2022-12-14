import sqlite3
import re


class Database():
    def __init__(self, dbname):
        self.conn = sqlite3.connect(dbname)
        self.cur = self.conn.cursor()
        print('Database conected!!!')

        self.cur.execute(
            '''
            CREATE TABLE IF NOT EXISTS clients(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            mail TEXT NOT NULL)
            '''
        )

        self.cur.execute(
            '''
            INSERT OR REPLACE INTO clients(id, name, mail)
            VALUES(1, 'HMT', 'info@hmt-automotive.com');
            '''
        )

    def insert_clients(self, name, email):
        name = self.name_clear_chars(name)
        self.cur.execute(
            '''
            INSERT INTO clients(id, name, mail)
            VALUES(Null, ?, ?)
            ''', (name, email)
        )
        self.conn.commit()

    def select_clients(self):
        self.cur.execute(
            '''
            SELECT * FROM clients
            ORDER BY name ASC
            '''
        )
        return self.cur.fetchall()

    def select_client_by_name(self, name):
        self.cur.execute(
            '''
            SELECT * FROM clients
            WHERE name = ?
            ''', (name,)
        )
        return self.cur.fetchall()

    def select_client_by_id(self, pk):
        self.cur.execute(
            '''
            SELECT * FROM clients
            WHERE id = ?
            ''', (pk,)
        )
        return self.cur.fetchall()

    def update_client(self, name, email, pk):
        name = self.name_clear_chars(name)
        self.cur.execute(
            '''
            UPDATE clients
            SET name=?, mail=?
            WHERE id = ?
            ''', (name, email, pk)
        )
        self.conn.commit()

    def delete_client(self, name):
        self.cur.execute(
            '''
            DELETE FROM clients
            WHERE name = ?
            ''', (name,)
        )
        self.conn.commit()

    def email_validation(self, mail):
        valid_email_regex = '^(\w|\.|\_|\-)+[@](\w|\_|\-|\.)+[.]\w{2,3}$'
        if mail != 'poczta' and not re.search(valid_email_regex, mail):
            return False
        return True

    def name_clear_chars(self, name):
        for c in '\/:*?"<>|':
            name = name.replace(c, '')
        return name
