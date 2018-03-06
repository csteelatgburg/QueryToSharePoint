# QueryToSharePoint
Uploads results of a MySQL query to a SharePoint list

This is a set of Python scripts that can be used to upload the results from a MySQL query to a SharePoint list on Office 365. I developed this script to make creation of SharePoint lists and keeping those lists in sync with our Quest K1000 database. Rather than sync the entire database, the script uses queries stored in .sql files that are executed against the database and results are uploaded to SharePoint. This allows users with little SQL experience to create reports in the wizard or other SQL query building tools and create the .sql files for the SharePoint lists that they need for their applications.

Query files are placed in the queries folder and when the script is run it will:

For each .sql file in the queries folder:
  If there isn't a List with the same name as the file (without the .sql extension), create the list
  If there is already a List, move on to the next file

After lists are checked against queries, each query is executed.
  For each row in the results, 
    If a matching item exists in the list, compare each column value with corresponding field value
      Update fields in List as necessary
    If a matching row does not exist in the List, add the row as a new item in the list
    If an item exists in the list that is not a row in the results, remove the item from the list


  
