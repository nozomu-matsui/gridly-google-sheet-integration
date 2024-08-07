from locale import normalize
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json, requests, csv, io
import gridly_api_handler

# define the scope
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('cred.json', scope)
client = gspread.authorize(creds)


def updateCells(viewId, sheetUniqueIdColumn, gridlyApiKey, spreadSheetName):
    sheetUniqueIdColumn = int(sheetUniqueIdColumn)
    global sheetTabs
    sheetTabs = client.open(spreadSheetName)._spreadsheets_get()
    steps = 100
    startNumber = 0
    untilNumber = 100   
    url = "https://api.gridly.com/v1/views/"+viewId+"/records?page=%7B%22offset%22%3A+"+str(startNumber)+"%2C+%22limit%22%3A+"+str(untilNumber)+"%7D"

    payload={}
    headers = {
    'Authorization': 'ApiKey ' + gridlyApiKey
    }    

    response = requests.request("GET", url, headers=headers, data=payload)
    records = json.loads(response.content)
    while len(response.content) >= steps-1:
        startNumber = untilNumber + 1
        untilNumber = untilNumber + steps
        url = "https://api.gridly.com/v1/views/"+viewId+"/records?page=%7B%22offset%22%3A+"+str(startNumber)+"%2C+%22limit%22%3A+"+str(untilNumber)+"%7D"
        response = requests.request("GET", url, headers=headers, data=payload)
        records += json.loads(response.content)
        #print(len(json.loads(response.content)))
    
    #print(len(records))

    lastPathName = ""
    lastTabId = 0
    sheetHeaders = []
    recordIds = []
    _sheet = ""
    updates = []
    sheet = client.open(spreadSheetName)
    sheets = sheet._spreadsheets_get()["sheets"]
    data = sheet.get_worksheet(0).get_all_records()
    keys = list(data[0].keys())
    sheetRecordIdAlias = keys[sheetUniqueIdColumn]
    #for record in response:
    lastPath = 9999
    for record in records:
        justChanged = False
        value = ""

        if(lastPathName != record["path"]):
            lastTabId = getTabIdByName(record["path"])
            lastPathName = record["path"]
            justChanged = True
            _sheet = client.open(spreadSheetName).get_worksheet(lastTabId)
            print("_sheet name: " + _sheet.title)
        if(justChanged):
            sheetAllRecord = _sheet.get_all_records()
            sheetHeaders.clear()
            recordIds.clear()
            for key in sheetAllRecord[int(sheetUniqueIdColumn)]:
                sheetHeaders.append(key)
            for sheetRecord in sheetAllRecord:
                recordIds.append(str(sheetRecord[sheetRecordIdAlias]))
        for cell in record["cells"]:
            normalized_sheet_headers = []
            for sHeader in sheetHeaders:
                normalized_sheet_headers.append(''.join(filter(str.isalnum, sHeader)))
            #print(normalized_sheet_headers)
            if "value" in cell:
                value = cell["value"]
            else:
                value = ""

            try:
                row = recordIds.index(record["id"])+1
                col = normalized_sheet_headers.index(cell["columnId"])
            except:
                continue

            if lastPath == 9999:
                lastPath = lastTabId
            if lastPath != lastTabId:
                _sheet = client.open(spreadSheetName).get_worksheet(lastPath)
                _sheet.batch_update(updates)
                updates = []
                lastPath = lastTabId

            updates.append({'range': 'R['+str(row)+']C['+str(col)+']:R['+str(row)+']C['+str(col)+']', 'values':[[ value ]]})
        

    _sheet = client.open(spreadSheetName).get_worksheet(lastTabId)
    _sheet.batch_update(updates)

    #print(newUpdates)



def getTabIdByName(tabName):
    for sheet in sheetTabs["sheets"]:
        if sheet["properties"]["title"] == tabName:
            return sheet["properties"]["index"]


def pullSheet(event, context):
    #event = json.loads(event)
    sheetUniqueIdColumn = event["sheetUniqueIdColumn"]
    synchColumns = event["synchColumns"]
    spreadSheetName = event["spreadSheetName"]
    viewId = event["viewId"]
    gridlyApiKey = event["gridlyApiKey"]
    getSheetAsCSV(spreadSheetName, viewId, gridlyApiKey, synchColumns, sheetUniqueIdColumn)



def pushSheet(event, context):
    #event = json.loads(event)
    sheetUniqueIdColumn = event["sheetUniqueIdColumn"]
    synchColumns = event["synchColumns"]
    spreadSheetName = event["spreadSheetName"]
    viewId = event["viewId"]
    gridlyApiKey = event["gridlyApiKey"]
    updateCells(viewId, sheetUniqueIdColumn, gridlyApiKey, spreadSheetName)



def getSheetAsCSV(spreadSheetName, viewId, gridlyApiKey, synchColumns, sheetUniqueIdColumn):
    sheet = client.open(spreadSheetName)
    sheets = sheet._spreadsheets_get()["sheets"]
    for i in range(0, len(sheets), 1):
        data = sheet.get_worksheet(i).get_all_records()
        headers = []
        for key in data[0]:
            headers.append(key)
        for item in data:
            item.update( {"_pathTag":sheets[i]["properties"]["title"]})

        gridly_api_handler.importCSV(json_to_csv(data, int(sheetUniqueIdColumn)), headers, viewId, gridlyApiKey, synchColumns, ExcludedColumnName)

    
def json_to_csv(jsonFile, sheetUniqueIdColumn):
    keys = list(jsonFile[0].keys())
    global ExcludedColumnName
    ExcludedColumnName = keys[sheetUniqueIdColumn]
    
    for rec in jsonFile:
        rec["_recordId"] = rec.pop(keys[sheetUniqueIdColumn])
    
    # Create an in-memory string buffer
    output = io.StringIO()
    
    # Create CSV writer for the string buffer
    writer = csv.DictWriter(output, fieldnames=jsonFile[0].keys(), delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    writer.writeheader()
    writer.writerows(jsonFile)
    
    # Get the CSV content as a string
    csv_content = output.getvalue()
    
    return csv_content

