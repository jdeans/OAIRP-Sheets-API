import csv
import cx_Oracle
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from string import ascii_uppercase


#Written in Python3 by Joshua Deans for presentation @ OAIRP 2017

def connect_sheets(keyfile, bookname, sheetname = 'Sheet1'):
    #connect to a workbook and sheet
    scope = ['https://spreadsheets.google.com/feeds']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(keyfile, scope)
    gc = gspread.authorize(credentials)
    wks = gc.open(bookname).worksheet(sheetname)

    return wks

def connect_banner(loc, port, name, username, pwd):
    #connect to banner, or another Oracle DB
    dsnstr = cx_Oracle.makedsn(loc, port, name)
    con = cx = cx_Oracle.connect(user = username, password = pwd, dsn = dsnstr)
    cur = con.cursor()

    return cur

def execute_query(query):
    #takes argument 'query' and passes it to Oracle via the cursor
    #returns the results as a list, which gspread likes
    cur.execute(query)

    return cur.fetchall()

def flatten(data):
    #flattens query results into a list
    return [item for siblist in data for item in sublist]

def get_range(start, unflattened_data):
    #takes a user defined starting point (e.g. A2) and calculates the
    #appropriate ending point (e.g. AB4567)
    return start + ':' + (ascii_uppercase[(len(unflattened_data)-1)//26]) +
                          str(len(unflattened_data)+1)

def main():

    query = ""
    
    chunksize = 10000
    for x in range(0, len(flattened_data), chunksize):
        wks.update_cells(cells[i:i+chunksize])

if __name__ == "__main__":
    main()
