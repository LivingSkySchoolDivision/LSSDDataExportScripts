# Edsby to SchoolLogic Marks import


# Usage

```
.\EdsbyMarksToSchoolLogicMarks.ps1 -inputfilename import.csv
```
Read and process the given Edsby Report Card Export file. 
Will read from your SchoolLogic/SIRS database, but will not actually write unless you add `-Commit $true`.

You should run this without `-Commit $true` (or with `-Commit $false`) first, to make sure there are no obvious data or import errors, before making actual changes to your database.

```
.\EdsbyMarksToSchoolLogicMarks.ps1 -inputfilename import.csv -Commit $true
```
Read, process, and then import the given Edsby Reprot Card Export file into your SchoolLogic/SIRS database.


# Advanced usage
```
   -ConfigFilePath
```
Specify the path to the config file. Config file should be an XML file that contains the database connection string. Example config file can be found below.
If not specified, it will default to __two__ levels above the current file (to match it's relative position in this git repository). ie: `../../config.xml`.

```
   -EmptyMarksLogPath empty.csv 
```
If there are any empty marks in your import file, they will be dumped into this CSV file. If this parameter is not specified, empty marks in the import file will be silently ignored.

```
   -OrphanedMarksLogPath orphans.csv 
```
If there are marks in your import file for classes/sections that do not exist in your SchoolLogic/SIRS database, they will be logged to this CSV file. If this parameter is not specified, these "orphaned" marks will be silently ignored.

```
   -ErrorLogPath errors.csv
```
If there are parsing errors when reading lines from the import file, the lines will be written to this CSV file. If this parameter is not specified, lines causing errors or exceptions will be ignored.


# Changes you need to make to your database
This script requires that you add __two__ fields to your `Marks` table in your SchoolLogic/SIRS database. These fields are non-intrusive, and are added to help you recover from a potential bad import by allowing you to easily refer to imported rows.

These fields should be added to the __Marks__ table.

## ImportTimestamp
Add a field named `ImportTimestamp`, of type `datetime` to the `Marks` table. This will be a timestamp of when you ran the script, which can help track errors or issues by helping you target specific rows that were added at a specific time.

## ImportBatchID
Add a field named `ImportBatchID`, of type `varchar(40)` to the `Marks` table. This will be a unique ID hash value for each time the import script is run. If there is a mistake or error, this allows you to easily target all records added by a specific script run in order to update or remove them.

# Example config file

```
<?xml version="1.0" encoding="utf-8" ?>
<Settings>
  <SchoolLogic>
    <ConnectionStringRW>data source=SERVERNAME;initial catalog=DATABASENAME;user id=USERNAME;password=PASSWORD;Trusted_Connection=false</ConnectionStringRW><!-- Read/Write-->
  </SchoolLogic>  
</Settings>

```
