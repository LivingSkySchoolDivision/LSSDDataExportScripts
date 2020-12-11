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
-ConfigFilePath config.xml
```
Specify the path to the config file. Config file should be an XML file that contains the database connection string. Example config file can be found below.
If not specified, it will default to __two__ levels above the current file (to match it's relative position in this git repository). ie: `../../config.xml`.

```
-EmptyMarksLogPath empty.csv 
```
If there are any empty marks in your import file, they will be dumped into this CSV file. If this parameter is not specified, empty marks in the import file will be silently ignored. If there are no empty marks, no file will be created.

```
-OrphanedMarksLogPath orphans.csv 
```
If there are marks in your import file for classes/sections that do not exist in your SchoolLogic/SIRS database, they will be logged to this CSV file. If this parameter is not specified, these "orphaned" marks will be silently ignored. If there are no orphaned marks, no file will be created.

```
-ErrorLogPath errors.csv
```
If there are parsing errors when reading lines from the import file, the lines will be written to this CSV file. If this parameter is not specified, lines causing errors or exceptions will be ignored. If there are no errors, no file will be created.

# Example

```
.\EdsbyMarksToSchoolLogicMarks.ps1 -ConfigFilePath config.xml -inputfilename import.csv -EmptyMarksLogPath empty.csv -OrphanedMarksLogPath orphans.csv -ErrorLogPath errors.csv -Commit $false
```
The above example imports a file and will identify any errors either in the console, or in one of three potential CSV files that are created.

```
.\EdsbyMarksToSchoolLogicMarks.ps1 -ConfigFilePath config.xml -inputfilename import.csv -Commit $true
```
Having dealt with all of the issues, the above example will import the file into the database.

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
