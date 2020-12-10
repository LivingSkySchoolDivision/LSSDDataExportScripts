# Edsby to SchoolLogic Marks import


# Usage

```
.\EdsbyMarksToSchoolLogicMarks.ps1 
   -inputfilename import.csv 
   -Commit $true
```

# Advanced usage
```
   -ConfigFilePath
```
Specify the path to the config file. Config file should be an XML file that contains the database connection string. Example config file can be found below.
```
   -EmptyMarksLogPath empty.csv 
```
```
   -OrphanedMarksLogPath orphans.csv 
```
```
   -ErrorLogPath errors.csv
```

# Changes you need to make to your database
 


# Example config file

```
<?xml version="1.0" encoding="utf-8" ?>
<Settings>
  <SchoolLogic>
    <ConnectionStringRW>data source=SERVERNAME;initial catalog=DATABASENAME;user id=USERNAME;password=PASSWORD;Trusted_Connection=false</ConnectionStringRW><!-- Read/Write-->
  </SchoolLogic>  
</Settings>

```