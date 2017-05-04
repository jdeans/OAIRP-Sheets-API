import cx_Oracle
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from string import ascii_uppercase

#connect to Oracle DB
dsnstr = cx_Oracle.makedsn("banner.university.edu", "port", "dsn name")
con = cx = cx_Oracle.connect(user = "username", password = "password", dsn = dsnstr)
cur = con.cursor()

#execute a query
cur.execute("Your query here! Exclude semi-colon")
data = cur.fetchall()

#close cursor and connection
cur.close()
con.close()

#connect to sheets API
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('API json keyfile location', scope)
gc = gspread.authorize(credentials)
wks = gc.open("my_workbook").worksheet("my_sheet")

#you must tell the sheets API which cells to select for updating
#this script always starts in cell A2 and then selects the appropriate 
#column identifier and row number
#we then feed this cell range into an object "cells"
if len(data[0]) <= 26:
    cell_range = ('A2:' + ascii_uppercase[len(data[0])-1]+str(len(data)+1))
else:
    cell_range = ('A2:' + ascii_uppercase[len(data[0])//26-1] +
                      ascii_uppercase[len(data[0])%26-1] +
                      str(len(data)+1))
print ('Selected cell range: %s' %(cell_range))
cells = wks.range(cell_range)

#flatten data from an array into a list
flattened_data = [item for sublist in data for item in sublist]

#clean and transform data
#This is just an example of some of the cleaning I did 
#to showcase why this process is advantageous
colleges = {'Arts and Sciences':'Arts & Sciences', 'Business Administration':'Business'}
flattened_data = [colleges[x] if x in colleges else x for x in flattened_data]
flattened_data = ['' if x == None else x for x in flattened_data]

print ('Updating %d rows' % (len(data)))

#feed the flattened data into the cells object
for x in range(len(flattened_data)):
    cells[x].value = flattened_data[x]

#the sheets API limits the number of cells we can update at once to 40k
#to be safe, I update in batches of 10k
chunksize = 10000
for i in range(0, len(flattened_data), chunksize):
    wks.update_cells(cells[i:i+chunksize])

#the script above pastes data into a specific cell range
#the script below can be used to append data to the bottom of
#your worksheet instead. It is a simpler, but much slower process
"""
for result in cur:
    wks.append_row(result)
"""
