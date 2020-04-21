# LSSD Third Party Integration Scripts

We use these scripts to create export files of our student and staff data, in order to integrate our systems with other third-party systems. 

These scripts only create the export/integration files. We use seperate scripts to ship the integration files to their respective third party systems.
 

# Requirements

* These scripts were written and tested using PowerShell version 5.1.
* These scripts do not require any additional PowerShell modules to be installed.
* These scripts were written and tested using Windows 10 and Windows Server 2016, and are not guaranteed to work on Linux or in Azure Cloud Shell.
* These scripts may rely on specific fields in our environment, which might not exist in your databases.

# Command line arguments

All scripts support the following standard arguments:

```
-OutputFileName <filename>
```
The file name (with path) of the file to output. Normally this is a CSV file, but may be another format based on the script. In many cases, the file name must match exactly what the third party system expects. Check the third party system's documentation for more information.

Paths are supported.

Required.

```
-ConfigFilePath <filename>
```
The full path to the configuration file (see below for more information regarding this file). 

Paths are supported.

Defaults to `config.xml` in the folder *above* the folder the script is stored in.

# Configuration file

The scripts in this repository require a database connection string in order to communicate with your database server.

This file must be in xml format, and must follow a specific format.

Example configuration file:
```xml
<?xml version="1.0" encoding="utf-8" ?>
<Settings>
  <SchoolLogic>
    <ConnectionString>data source=SERVERNAME;initial catalog=DATABASENAME;user id=USERNAME;password=PASSWORD;Trusted_Connection=false</ConnectionString>
  </SchoolLogic>
  <Navision>
    <ConnectionString>data source=SERVERNAME;initial catalog=DATABASENAME;Trusted_Connection=true</ConnectionString>
  </Navision>
</Settings>

```

An example config file is provided. Simply create a copy of this file named "config.xml", and edit it to contain your own database connection string(s).

Some scripts may require additional connection strings, as they may need to communicate with different databases or multiple databases.

If you require assistance figuring out what connection string you need, the following third-party sites may help you:
* https://www.connectionstrings.com/
* https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms722656(v=vs.85)