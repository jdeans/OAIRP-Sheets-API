import cx_Oracle
from string import ascii_uppercase
from pprint import pprint

from oauth2client import file, client, tools
from oauth2client.service_account import ServiceAccountCredentials
from apiclient.discovery import build
from httplib2 import Http
from googleapiclient import discovery

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

#Connect to the sheets API. secrets.json is your keyfile
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('secrets.json', scope)
http_auth = credentials.authorize(Http())


service = discovery.build('sheets', 'v4', http=credentials.authorize(Http()))

#select spreadsheet to paste data into
spreadsheet_id = 'Your sheet ID here!'

#select cell range to update based on data. 
if len(data[0]) <= 26:
    range_ = ('A2:' + ascii_uppercase[len(data[0])-1]+str(len(data)+1))
else:
    range_ = ('A2:' + ascii_uppercase[len(data[0])//26-1] +
                      ascii_uppercase[len(data[0])%26-1] +
                      str(len(data)+1))
print ('Selected cell range: %s' %(cell_range))

#if you are using a script such as this, you should keep this value as is
#see API documentation for other options and their uses
value_input_option = 'RAW'

#this is the meat of the data upload
#majorDimension can be set to rows or columns, usually rows is what you want
value_range_body = {
    "range": range_,
    "majorDimension": "ROWS",
    "values": data
}

#below is the request to the API. We simply feed it our variables defined above
#it is helpful to execute and save the server response in a variable
request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, 
                                                 range=range_, 
                                                 valueInputOption=value_input_option, 
                                                 body=value_range_body)
response = request.execute()
pprint(response)
