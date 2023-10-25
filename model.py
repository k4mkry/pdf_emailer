import sqlite3
import re


class Database:
    def __init__(self, dbname):
        self.conn = sqlite3.connect(dbname)
        self.cur = self.conn.cursor()
        print("Database conected!!!")

        self.cur.execute(
            """
            CREATE TABLE IF NOT EXISTS clients(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            mail TEXT NOT NULL)
            """
        )
        self.conn.commit()

        self.cur.execute(
            """
            CREATE TABLE IF NOT EXISTS reports(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nr_faktury TEXT NOT NULL,
            data TEXT NOT NULL,
            kontrahent TEXT NOT NULL)
            """
        )
        self.conn.commit()

    def add_report(self, nr_faktury, data, kontrahent):
        # nr_faktury, data, konrahent
        self.cur.execute(
            """
            INSERT INTO reports(id, nr_faktury, data, kontrahent)
            VALUES(Null, ?, ?, ?)
            """,
            (nr_faktury, data, kontrahent),
        )
        self.conn.commit()

    def get_report(self):
        self.cur.execute(
            """
            SELECT * FROM reports
            ORDER BY data ASC
            """
        )
        return self.cur.fetchall()

    def insert_clients(self, name, email):
        name = self.name_clear_chars(name)
        self.cur.execute(
            """
            INSERT INTO clients(id, name, mail)
            VALUES(Null, ?, ?)
            """,
            (name, email),
        )
        self.conn.commit()

    def select_clients(self):
        self.cur.execute(
            """
            SELECT * FROM clients
            ORDER BY name ASC
            """
        )
        return self.cur.fetchall()

    def select_client_by_name(self, name):
        self.cur.execute(
            """
            SELECT * FROM clients
            WHERE name = ?
            """,
            (name,),
        )
        return self.cur.fetchall()

    def select_client_by_id(self, pk):
        self.cur.execute(
            """
            SELECT * FROM clients
            WHERE id = ?
            """,
            (pk,),
        )
        return self.cur.fetchall()

    def update_client(self, name, email, pk):
        name = self.name_clear_chars(name)
        self.cur.execute(
            """
            UPDATE clients
            SET name=?, mail=?
            WHERE id = ?
            """,
            (name, email, pk),
        )
        self.conn.commit()

    def delete_client(self, name):
        self.cur.execute(
            """
            DELETE FROM clients
            WHERE name = ?
            """,
            (name,),
        )
        self.conn.commit()

    def email_validation(self, mail):
        valid_email_regex = "^(\w|\.|\_|\-)+[@](\w|\_|\-|\.)+[.]\w{2,3}$"
        if mail != "poczta" and not re.search(valid_email_regex, mail):
            return False
        return True

    def name_clear_chars(self, name):
        for c in '\/:*?"<>|':
            name = name.replace(c, "")
        return name
