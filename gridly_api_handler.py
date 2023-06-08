import requests
from requests_toolbelt.multipart.encoder import MultipartEncoder
import json
import os
import time



viewID = ""


def importCSV(mycsv, sheetHeaders, _viewId, _gridlyApiKEy, synchColumns, _ExcludedColumnName):
    ExcludedColumnName = _ExcludedColumnName
    global viewId
    global gridlyApiKEy
    gridlyApiKEy = _gridlyApiKEy
    viewId = _viewId

    refreshView()

    while(view["gridStatus"] != "active"):
        time.sleep(30)
        refreshView()

    if synchColumns == "true":
        synchHeaders(sheetHeaders, ExcludedColumnName)
        
    url = "https://api.gridly.com/v1/views/" + viewId + "/import"

    mycsv = str.encode(mycsv)

    mp_encoder = MultipartEncoder(
        fields={
        'file': ('addresses.csv',mycsv,'text/csv')
        }
        )

    headers = {
    'Authorization': 'ApiKey ' + gridlyApiKEy,
    'Content-Type': mp_encoder.content_type
    }
    importResponse = requests.request("POST", url, headers=headers, data=mp_encoder)
    #print(importResponse.content)
    time.sleep(5)

def refreshView():
    url = "https://api.gridly.com/v1/views/" + viewId

    payload={}
    headers = {
    'Authorization': 'ApiKey ' + gridlyApiKEy
    }
    global view
    view = json.loads(requests.request("GET", url, headers=headers, data=payload).content)

def getGridlyHeaders():
    columnNames = []
    try:
        for column in view["columns"]:
            if "name" in column:
                columnNames.append(column["name"])
    except Exception as e:
        return None
    return columnNames


def synchHeaders(sheetHeaders, ExcludedColumnName):
    #print(ExcludedColumnName)
    gridlyHeaders = getGridlyHeaders()
    for sheetheader in sheetHeaders:
        if sheetheader not in gridlyHeaders and sheetheader != "_recordId" and sheetheader != "_pathTag" and sheetheader != ExcludedColumnName:
            createGridlyHeader(sheetheader)
    refreshView()


def createGridlyHeader(headerName):
    refreshView()
    if "gridStatus" in view:
        while(view["gridStatus"] != "active"):
            time.sleep(30)
            refreshView()

    url = "https://api.gridly.com/v1/views/" + viewId + "/columns"
    id = ''.join(filter(str.isalnum, headerName))
    payload = json.dumps({
    "name": headerName,
    "type": "multipleLines",
    "id": id
    })
    headers = {
    'Authorization': 'ApiKey ' + gridlyApiKEy,
    'Content-Type': 'application/json'
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    #print(response.text)
