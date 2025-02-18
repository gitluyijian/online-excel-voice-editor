from flask_login import UserMixin
import sqlite3
from werkzeug.security import generate_password_hash, check_password_hash

class User(UserMixin):
    def __init__(self, id, username, password_hash):
        self.id = id
        self.username = username
        self.password_hash = password_hash

    @staticmethod
    def get(user_id):
        db = sqlite3.connect('users.db')
        cursor = db.cursor()
        cursor.execute('SELECT * FROM users WHERE id = ?', (user_id,))
        user = cursor.fetchone()
        db.close()
        if user:
            return User(user[0], user[1], user[2])
        return None

    @staticmethod
    def get_by_username(username):
        db = sqlite3.connect('users.db')
        cursor = db.cursor()
        cursor.execute('SELECT * FROM users WHERE username = ?', (username,))
        user = cursor.fetchone()
        db.close()
        if user:
            return User(user[0], user[1], user[2])
        return None

    @staticmethod
    def authenticate(username, password):
        db = sqlite3.connect('users.db')
        cursor = db.cursor()
        cursor.execute('SELECT * FROM users WHERE username = ?', (username,))
        user = cursor.fetchone()
        db.close()
        if user and check_password_hash(user[2], password):
            return User(user[0], user[1], user[2])
        return None

def init_db():
    db = sqlite3.connect('users.db')
    cursor = db.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL
        )
    ''')
    db.commit()
    db.close() 