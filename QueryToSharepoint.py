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
from MySQLfun import *
from SPfun import *



def json_serial(obj):
    """JSON serializer for objects not serializable by default json code"""

    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    raise TypeError ("Type %s not serializable" % type(obj))

def queriesGet():
    files = os.listdir("queries")
    for file in files:
        if not file.endswith(".sql"):
            #it isn't a query, get rid of it
            del file

    return files

def checkQueriesAndLists(lists, queries):
    errorState = 0
    listTitles = []
    for list in lists:
        #print(list['Id'], list['Title'])
        listTitles.append(list['Title'])
    for query in queries:
        if query.replace(".sql", "") in listTitles:
            print(query + " exists as a list, will update")
        else:
            print(query + " does not exist as a list, will create")
            queryColumns = MySQLColumnTypes(query)
            #print(queryColumns)
            #check columns for a UID column
            UIDColumn = 0
            for column in queryColumns:
                if column[0] == "UID":
                    UIDColumn += 1
                elif column[0] == "ID":
                    #reserved field name in Sharepoint, output an error
                    print("ERROR: cannot create a field in Sharepoint named ID.")
                    print("Please update your query and try again.")
                    errorState = 1

                elif " " in column[0]:
                    #can't handle spaces in column names
                    print("ERROR: cannot create a field in Sharepoint with a space in the name.")
                    print("Please update your query and try again.")
                    errorState = 1

            if UIDColumn == 0:
                print("ERROR: Your query does not contain a UID column.")
                print("Script relies on a column named UID which contains a unique value.")
                print("Please update your query and try again.")
                errorState = 1

            if UIDColumn > 1:
                print("ERROR: Your query contains more than one UID column.")
                print("Script relies on one column named UID which contains a unique value.")
                print("Please update your query and try again.")
                errorState = 1

            if errorState == 0:
                newList = SPListCreate(query.replace(".sql", ""), authheader)
                print(newList)
                #Create field in Sharepoint for each column in query
                for column in queryColumns:
                    if column[1] in dateTimeTypes:
                        print(column[0] + " will be a datetime column")
                        createField = SPListCreateField(newList, column[0], 4, authheader)
                    elif column[1] in integerTypes:
                        print(column[0] + " will be an integer column")
                        createField = SPListCreateField(newList, column[0], 1, authheader)
                    elif column[1] in numberTypes:
                        print(column[0] + " will be a number column")
                        createField = SPListCreateField(newList, column[0], 9, authheader)
                    elif column[0][:7] == "PERSON-":
                        print(column[0] + " will be a person column")
                        createField = SPListCreateField(newList, column[0], 20, authheader)

                    else:
                        print(column[0] + " will be a text column")
                        createField = SPListCreateField(newList, column[0], 2, authheader)
    return errorState;

def updateSPItemsFromMyRows(listName, SPItems, rows, types):

    global updated
    global deleted
    for item in SPItems:
        if item['UID'] in rows:
            fieldUpdated = 0
            #print("Checking fields for UID " + str(item['UID']))
            fields = json.loads(json.dumps(rows[item['UID']], default=json_serial))
            #print(fields)
            for field in fields:
                #print(field)
                #print(field, types[field], item[field])
                if types[field] == 10 and item[field] is not None:
                    #date only field
                    from_zone = tz.tzutc()
                    to_zone = tz.tzlocal()
                    SPDate = datetime.strptime(item[field].replace('Z', '+0000'), '%Y-%m-%dT%H:%M:%S%z').astimezone(to_zone).strftime('%Y-%m-%d')
                    MySQLDate = datetime.strptime(fields[field], '%Y-%m-%d').strftime('%Y-%m-%d')
                    if SPDate != MySQLDate:
                        print("Difference found for item UID: " + str(item['UID']))
                        print("updating " + str(field) + " field from " + SPDate + " to " + MySQLDate)
                        updatestatus = SPUpdateItemField(listName, item['ID'], field, fields[field], authheader)
                        print(updatestatus.decode("utf-8"))
                        fieldUpdated += 1
                elif types[field] == 7 or types[field] == 12:
                    #date time field, correct for timezone
                    from_zone = tz.tzutc()
                    to_zone = tz.tzlocal()
                    SPDateTime = datetime.strptime(item[field].replace('Z', '+0000'), '%Y-%m-%dT%H:%M:%S%z').astimezone(to_zone).strftime('%Y-%m-%dT%H:%M:%S')
                    MySQLDateTime = datetime.strptime(fields[field], '%Y-%m-%dT%H:%M:%S').strftime('%Y-%m-%dT%H:%M:%S')
                    #print(item[field])
                    #print(fields[field])
                    if SPDateTime != MySQLDateTime:
                        print("Difference found for item UID: " + str(item['UID']))
                        print("updating " + str(field) + " field from " + SPDateTime + " to " + MySQLDateTime)
                        updatestatus = SPUpdateItemField(listName, item['ID'], field, fields[field], authheader)
                        print(updatestatus.decode("utf-8"))
                        fieldUpdated += 1
                elif str(field)[:7] == "PERSON-":

                    #special field type of person.
                    #field is updated using different key
                    idKey = str(field).replace("-", "_x002d_") + "Id"
                    if item[idKey] is not None and item[idKey] != fields[field]:
                        print("Difference found for item UID: " + str(item['UID']))
                        print("updating " + str(field) + " field from " + str(item[idKey]) + " to " + str(fields[field]))
                        updatestatus = SPUpdateItemField(listName, item['ID'], idKey, str(fields[field]), authheader)
                        print(updatestatus.decode("utf-8"))
                        fieldUpdated += 1

                elif item[field] != fields[field] and item[field] is not None:
                    print("Difference found for item UID: " + str(item['UID']))
                    print("updating " + str(field) + " field from " + str(item[field]) + " to " + str(fields[field]))
                    updatestatus = SPUpdateItemField(listName, item['ID'], field, fields[field], authheader)
                    print(updatestatus.decode("utf-8"))
                    fieldUpdated += 1
                elif item[field] is None and fields[field] is not None and len(fields[field]) > 0:
                    #print(len(fields[field]))
                    print("Difference found for item UID: " + str(item['UID']))
                    print("updating " + str(field) + " field from " + str(item[field]) + " to " + str(fields[field]))
                    updatestatus = SPUpdateItemField(listName, item['ID'], field, fields[field], authheader)
                    print(updatestatus.decode("utf-8"))
                    fieldUpdated += 1


            del rows[item['UID']]
            if fieldUpdated > 0:
                updated += 1
        else:
            deleteItem = SPDelItem(listName, item, authheader)
            deleted += 1

    return rows

authheader=authenticate(config['client_id'], config['tenant_id'], config['client_secret'], config['resource_id'], config['sharepoint_url'])

#Get list of users on site
users = getUsers(authheader)

#create dictionary of users to search via email
siteUsers = {}
for user in users:
    #print(user['Id'], user['Email'])
    siteUsers[user['Email']] = user['Id']

updated = 0
deleted = 0
added = 0

#get lists from Sharepoint site
lists = SPLists(authheader)

#get queries from folder
queries = queriesGet()

#compare lists and queries and create lists if needed
comparisonStatus = checkQueriesAndLists(lists, queries)

if comparisonStatus != 0:
    print("There was an error, please check output")
else:
    print("Queries and lists match")
    #sync queries to lists
    for query in queries:
        updated = 0
        deleted = 0
        added = 0
        listName = query.replace(".sql", "")
        print("Processing " + listName)
        fh = open("queries/" + query, "r")
        sqlQuery = fh.read()
        MyRows = MySQLQuery(sqlQuery)

        queryColumns = MySQLColumnTypes(query)
        types = {}
        for column in queryColumns:
            types[column[0]] = column[1]

        #create new dict containing results from query
        #relies on report having a column named UID that is unique
        rows = {}
        for row in MyRows:
            for column in row.keys():
                if column[:7] == "PERSON-":
                    #special field type of person. Expects email address
                    #get ID for user in site based on given email address
                    if row[column] in siteUsers:
                        userID = siteUsers[row[column]]

                    else:
                        ensure = SPensureUser(str(row[column]), authheader)
                        #print(ensure)
                        if "odata.error" in ensure:
                            #can't find a user with this email address, going to move on
                            userID = 0
                            siteUsers[row[column]] = 0
                        else:
                            userID = ensure['Id']
                            siteUsers[row[column]] = userID
                    row[column] = userID
                    #print(userID)

            rows[row['UID']] = row
        print(str(len(rows)) + " rows in report")

        #get list based on query name
        SPItems = SPListItems(listName, authheader)
        print(str(len(SPItems)) + " rows in SP list")

        newRows = updateSPItemsFromMyRows(listName, SPItems, rows, types)

        print(str(updated) + " rows updated in Sharepoint")
        print(str(deleted) + " rows deleted in Sharepoint")
        print(str(len(newRows)) + " rows to be added to Sharepoint")
        if len(newRows) > 0:
            for row in newRows:
                #print(listName, newRows[row])
                print('.', end='')
                sys.stdout.flush()
                addItem = SPAddItem(listName, newRows[row], authheader)
                #print(addItem)
                added += 1

            print(str(added) + " rows have been added")
