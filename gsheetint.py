import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json, requests, csv
import gridly_api_handler

# define the scope
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('cred.json', scope)
client = gspread.authorize(creds)


def updateCells(viewId, sheetUniqueIdColumn, gridlyApiKey, spreadSheetName):
    global sheetTabs
    sheetTabs = client.open(spreadSheetName)._spreadsheets_get()
    url = "https://api.gridly.com/v1/views/"+viewId+"/records"

    payload={}
    headers = {
    'Authorization': 'ApiKey ' + gridlyApiKey
    }

    response = requests.request("GET", url, headers=headers, data=payload)

    lastPathName = ""
    lastTabId = 0
    sheetHeaders = []
    recordIds = []
    _sheet = ""
    #for record in response:
    for record in json.loads(response.content):
        justChanged = False
        updates = []

        if(lastPathName != record["path"]):
            lastTabId = getTabIdByName(record["path"])
            lastPathName = record["path"]
            justChanged = True
            _sheet = client.open(spreadSheetName).get_worksheet(lastTabId)
        if(justChanged):
            sheetAllRecord = _sheet.get_all_records()
            sheetHeaders.clear()
            recordIds.clear()
            for key in sheetAllRecord[int(sheetUniqueIdColumn)]:
                sheetHeaders.append(key)
            for sheetRecord in sheetAllRecord:
                recordIds.append(str(sheetRecord["_recordId"]))
        for cell in record["cells"]:

            if "value" not in cell:
                continue
            row = recordIds.index(record["id"])+1
            col = sheetHeaders.index(cell["columnId"])
            updates.append({'range': 'R['+str(row)+']C['+str(col)+']:R['+str(row)+']C['+str(col)+']', 'values':[[ cell["value"] ]]})


        _sheet.batch_update(updates)


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
    getSheetAsCSV(spreadSheetName, viewId, gridlyApiKey, synchColumns)
    


def pushSheet(event, context):
    #event = json.loads(event)
    sheetUniqueIdColumn = event["sheetUniqueIdColumn"]
    synchColumns = event["synchColumns"]
    spreadSheetName = event["spreadSheetName"]
    viewId = event["viewId"]
    gridlyApiKey = event["gridlyApiKey"]
    updateCells(viewId, sheetUniqueIdColumn, gridlyApiKey, spreadSheetName)



def getSheetAsCSV(spreadSheetName, viewId, gridlyApiKey, synchColumns):
    sheet = client.open(spreadSheetName)
    sheets = sheet._spreadsheets_get()["sheets"]
    for i in range(0, len(sheets), 1):
        data = sheet.get_worksheet(i).get_all_records()
        headers = []
        for key in data[0]:
            headers.append(key)
        for item in data:
            item.update( {"_pathTag":sheets[i]["properties"]["title"]})
        gridly_api_handler.importCSV(json_to_csv(data), headers, viewId, gridlyApiKey, synchColumns)

    
def json_to_csv(jsonFile):
    csvData = ""
    lines = 0
    for record in jsonFile:
        if lines == 0:
            for key in jsonFile[0].keys():
                csvData += key + "\t"
                lines = 1
            csvData = csvData[:-1]
            csvData += "\n"  
        for rec in record.values():      
            csvData += str(rec) + "\t"
        csvData = csvData[:-1]
        csvData += "\n"


        

    return csvData


# 
#getSheetAsCSV("Gridly Google Sheets integration", "VIEWID", "APIKEY", "true")
#pullSheet("{\r\n\"spreadSheetName\":\"SHEETNAME\",\r\n\"gridlyApiKey\":\"APIKEY\",\r\n\"viewId\":\"VIEWID\",\r\n\"synchColumns\":\"true\",\r\n\"sheetUniqueIdColumn\":0\r\n}", "")
#pushSheet("{\"spreadSheetName\": \"SHEETNAME\", \"gridlyApiKey\": \"APIKEY\", \"viewId\": \"VIEWID\", \"synchColumns\": \"true\", \"sheetUniqueIdColumn\": \"0\"}", "")

