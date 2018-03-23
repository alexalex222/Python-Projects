import sqlite3
import pandas as pd


def query_database(sql_query_statement):
    conn = sqlite3.connect('/Users/kuilinchen/anaconda3/bin/testDB.db')
    df = pd.read_sql_query(sql_query_statement, conn).to_string()
    conn.close()
    return df

def add(a,b):
    return a+b


if __name__ == '__main__':
    sql_statement = "SELECT * FROM stocks"
    print(query_database(sql_statement))
