import http.client
import json
import platform
import getpass
import urllib.request
import urllib.parse
import sys
import pymysql.cursors
from datetime import date, datetime, timedelta
import os
from dateutil import tz
from config import *

dateTimeTypes = [7, 10, 11, 12, 13, 14]
numberTypes = [0, 1, 2, 4, 5, 16, 246, 247]
integerTypes = [3, 8, 9]

def json_serial(obj):
    """JSON serializer for objects not serializable by default json code"""

    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    raise TypeError ("Type %s not serializable" % type(obj))

def SPLists(authheader):
    #create connection to submit data to sharepoint list
    #create headers
    headers = {
        'authorization': authheader,
        'accept': "application/json",
        'x-requestdigest': "form digest value",
        'content-type': "application/json",
        'cache-control': "no-cache"
        }

    submitconn = http.client.HTTPSConnection(config['sharepoint_url'])
    submiturl = config['site_path'] + "/_api/web/lists"
    submitconn.request("GET", submiturl, headers=headers)
    submitres = submitconn.getresponse()
    submitdata = submitres.read()
    #print(submitdata.decode("utf-8"))
    return json.loads(submitdata.decode("utf-8"))['value'];


def SPListCreate(listName, authheader):
    headers = {
        'authorization': authheader,
        'accept': "application/json; odata=verbose",
        'x-requestdigest': "form digest value",
        'content-type': "application/json; odata=verbose",
        'cache-control': "no-cache",
        }
    submitconn = http.client.HTTPSConnection(config['sharepoint_url'])
    submiturl = config['site_path'] + "/_api/web/lists"
    payload = "{ '__metadata': { 'type': 'SP.List' }, 'AllowContentTypes': true, 'BaseTemplate': 100, "
    payload += "'ContentTypesEnabled': true, 'Description': '" + listName + "', 'Title': '" + listName + "' }"
    #print(payload)
    submitconn.request("POST", submiturl, payload, headers)
    submitres = submitconn.getresponse()
    submitdata = submitres.read()
    result = json.loads(submitdata.decode("utf-8"))
    listURI = result['d']['__metadata']['uri']

    #print the status of the response
    #print(submitres.status)
    #print(submitdata.decode("utf-8"))

    return listURI

def SPListCreateField(listURI, FieldName, FieldType, authheader):
    headers = {
        'authorization': authheader,
        'accept': "application/json; odata=verbose",
        'x-requestdigest': "form digest value",
        'content-type': "application/json; odata=verbose",
        'cache-control': "no-cache"
        }
    submitconn = http.client.HTTPSConnection(config['sharepoint_url'])
    submitURL = listURI + "/Fields"
    payload = "{ '__metadata': { 'type': 'SP.Field' }, "
    payload += "'Title': '" + FieldName + "', 'FieldTypeKind': " + str(FieldType) + ", "
    payload += "'Required': false, 'EnforceUniqueValues': false, 'StaticName': '" + FieldName + "' }"
    #print(payload)
    submitconn.request("POST", submitURL, payload, headers)
    submitres = submitconn.getresponse()
    submitdata = submitres.read()

    #print the status of the response
    #print(submitres.status)
    #print(submitdata.decode("utf-8"))

    return submitdata

def SPListItems(listName, authheader):
    #create connection to submit data to sharepoint list
    #create headers
    headers = {
        'authorization': authheader,
        'accept': "application/json",
        'x-requestdigest': "form digest value",
        'content-type': "application/json",
        'cache-control': "no-cache"
        }

    submitconn = http.client.HTTPSConnection(config['sharepoint_url'])
    submiturl = config['site_path'] + "/_api/web/lists/GetByTitle('" + urllib.parse.quote(listName) + "')/Items?$top=5000"
    submitconn.request("GET", submiturl, headers=headers)
    submitres = submitconn.getresponse()
    submitdata = submitres.read()
    #print(submitdata.decode("utf-8"))
    return json.loads(submitdata.decode("utf-8"))['value'];

def SPDelItem(listName, SPItem, authheader):
    headers = {
        'authorization': authheader,
        'accept': "application/json; odata=verbose",
        'x-requestdigest': "form digest value",
        'content-type': "application/json; odata=verbose",
        'cache-control': "no-cache",
        "IF-MATCH": "*",
        'X-HTTP-Method':'DELETE'
        }
    submitconn = http.client.HTTPSConnection(config['sharepoint_url'])
    submiturl = config['site_path'] + "/_api/web/lists/GetByTitle('" + urllib.parse.quote(listName) + "')/items(" + str(SPItem['ID']) + ")"
    payload = ""
    submitconn.request("POST", submiturl, payload, headers)
    submitres = submitconn.getresponse()
    submitdata = submitres.read()

    #print the status of the response
    #print(submitres.status)
    #print(submitdata.decode("utf-8"))
    return submitdata

def SPUpdateItemField(listName, ItemID, Field, Value, authheader):
    #Updates the given item field with new value
    headers = {
        'authorization': authheader,
        'accept': "application/json; odata=verbose",
        'x-requestdigest': "form digest value",
        'content-type': "application/json",
        'cache-control': "no-cache",
        "IF-MATCH": "*",
        'X-HTTP-Method':'MERGE'
        }
    submitconn = http.client.HTTPSConnection(config['sharepoint_url'])
    submiturl = config['site_path'] + "/_api/web/lists/GetByTitle('" + urllib.parse.quote(listName) + "')/items(" + str(ItemID) + ")"
    fields = '"' + str(Field) + '": "' + str(Value).replace("\\", "\\\\") + '"'
    payload = "{" + fields + " }"
    #print(payload)
    try:
        submitconn.request("POST", submiturl, payload, headers)
        submitres = submitconn.getresponse()
        submitdata = submitres.read()

    except ConnectionRefusedError:
        print("Error connecting to update, will try again with next record.")


    #print the status of the response
    #print(submitres.status)
    #print(submitdata.decode("utf-8"))

    return submitdata

def SPAddItem(listName, MySQLItem, authheader):
    for key in MySQLItem.keys():
        if key[:7] == "PERSON-":
            newKey = key.replace("-", "_x002d_") + "Id"
            MySQLItem[newKey] = MySQLItem.pop(key)


    headers = {
        'authorization': authheader,
        'accept': "application/json; odata=verbose",
        'x-requestdigest': "form digest value",
        'content-type': "application/json",
        'cache-control': "no-cache",
        }
    submitconn = http.client.HTTPSConnection(config['sharepoint_url'])
    submiturl = config['site_path'] + "/_api/web/lists/GetByTitle('" + urllib.parse.quote(listName) + "')/items"
    fields = str(json.dumps(MySQLItem, default=json_serial))


    #print(fields)
    #print(payload)
    submitconn.request("POST", submiturl, fields, headers)
    submitres = submitconn.getresponse()
    submitdata = submitres.read()

    #print the status of the response
    #print(submitres.status)
    #print(submitdata.decode("utf-8"))

    return submitdata

def SPensureUser(email, authheader):
    headers = {
        'authorization': authheader,
        'accept': "application/json",
        'x-requestdigest': "form digest value",
        'content-type': "application/json",
        'cache-control': "no-cache"
        }

    submitconn = http.client.HTTPSConnection(config['sharepoint_url'])
    submiturl = config['site_path'] + "/_api/web/ensureuser"
    data = "{ 'logonName': 'i:0#.f|membership|" + email + "' }"
    submitconn.request("POST", submiturl, data, headers)
    submitres = submitconn.getresponse()
    submitdata = submitres.read()
    #print(submitdata.decode("utf-8"))
    return json.loads(submitdata.decode("utf-8"))

def getUsers(authheader):
    headers = {
        'authorization': authheader,
        'accept': "application/json",
        'x-requestdigest': "form digest value",
        'content-type': "application/json",
        'cache-control': "no-cache"
        }

    submitconn = http.client.HTTPSConnection(config['sharepoint_url'])
    submiturl = config['site_path'] + "/_api/web/siteusers?"
    #data = "{ 'logonName': 'i:0#.f|membership|" + email + "' }"
    submitconn.request("GET", submiturl, headers=headers)
    submitres = submitconn.getresponse()
    submitdata = submitres.read()
    #print(submitdata.decode("utf-8"))
    return json.loads(submitdata.decode("utf-8"))['value']

def authenticate(client_id, tenant_id,client_secret, resource_id, sharepoint_url):
    #print(config['client_id'])
    #create body for getting access_token
    body = "grant_type=client_credentials"
    body += "&client_id=" + client_id + "@" + tenant_id
    body += "&client_secret=" + client_secret
    body += "&resource=" + resource_id + "/" + sharepoint_url + "@" + tenant_id

    #print(body)

    #create headers for getting access_token
    headers = {
        'content-type': "application/x-www-form-urlencoded",
        'cache-control': "no-cache"
        }

    url = "/" + tenant_id + "/tokens/OAuth/2"

    #create connection to get authentication token
    conn = http.client.HTTPSConnection("accounts.accesscontrol.windows.net")
    conn.request("POST", url, body, headers)

    #place response in json
    res = conn.getresponse()
    data = res.read()
    resp = json.loads(data.decode("utf-8"))

    #print(resp)

    #get access_token from response body
    token = resp['access_token']

    #set variable for the authentication header
    authheader = "Bearer " + str(token)
    #print(authheader);

    #print( headers)
    return authheader;
