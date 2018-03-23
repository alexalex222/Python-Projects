import sqlite3
import pandas as pd


print(__name__)


def query_database(sql_query_statement):
    conn = sqlite3.connect('/Users/kuilinchen/anaconda3/bin/testDB.db')
    query_result_text = pd.read_sql_query(sql_query_statement, conn).to_string()
    conn.close()
    return query_result_text

def add(a,b):
    return a+b


if __name__ == '__main__':
    sql_statement = "SELECT * FROM stocks"
    result = query_database(sql_statement)
    print(result)
