import mysql.connector
import sqlite3
import os
def connect_db():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="", 
        database="conge"
    )
# def connect_db():
#     db_path = os.path.abspath("gestion_rh.db") 
#     return sqlite3.connect(db_path)