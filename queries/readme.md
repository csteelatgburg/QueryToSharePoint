Query files should be placed in this folder with the .sql file extension. Lists will be created using the file name, without the extension.

Query requirements:
* All queries must include a column named UID that contains a unique value.
This is used to correlate query rows with list items.
* Queries cannot contain a column named ID. It is reserved by SharePoint.
* Column names with spaces are not currently supported.
* SharePoint fields will be created based on the following column types
  * DECIMAL = 0 -> number
  * TINY = 1 -> number
  * SHORT = 2 -> number
  * LONG = 3 -> integer
  * FLOAT = 4 -> number
  * DOUBLE = 5 -> number
  * NULL = 6 -> string
  * TIMESTAMP = 7 -> dateTime
  * LONGLONG = 8 -> integer
  * INT24 = 9 -> integer
  * DATE = 10 -> dateTime
  * TIME = 11 -> dateTime
  * DATETIME = 12 -> dateTime
  * YEAR = 13 -> dateTime
  * NEWDATE = 14 -> dateTime
  * VARCHAR = 15 -> string
  * BIT = 16 -> number
  * JSON = 245 -> string
  * NEWDECIMAL = 246 -> number
  * ENUM = 247 -> number
  * SET = 248 -> string
  * TINY_BLOB = 249 -> string
  * MEDIUM_BLOB = 250 -> string
  * LONG_BLOB = 251 -> string
  * BLOB = 252 -> string
  * VAR_STRING = 253 -> string
  * STRING = 254 -> string
  * GEOMETRY = 255 -> string
  * CHAR = TINY
  * INTERVAL = ENUM
  * See [PyMySQL constants](https://github.com/PyMySQL/PyMySQL/blob/master/pymysql/constants/FIELD_TYPE.py) for reference.
* Columns beginning with PERSON- will be created as a Person field. These columns should contain the email address of users on the Office 365 tenant. When data is added to these fields the user will be ensured on the SharePoint site and the field will contain a reference to the user.
