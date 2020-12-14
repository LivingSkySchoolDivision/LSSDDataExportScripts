# LSSD Third Party Integration Scripts

We use these scripts to create export files of our student and staff data, in order to integrate our systems with other third-party systems. 

These scripts are not designed to be a drop in solution for data syncing, but we hope that they might be a good starting point for others.

These scripts only create the export/integration files. We use seperate scripts to ship the integration files to their respective third party systems.
 

# Requirements

* These scripts were written and tested using PowerShell version 5.1.
* These scripts do not require any additional PowerShell modules to be installed.
* These scripts were written and tested using Windows 10, Windows Server 2016, and Windows Server 2019. They are not guaranteed to work on Linux or in Azure Cloud Shell.
* These scripts may rely on specific fields in our environment, which might not exist in your databases. Some scripts may require modifications to your database.

# Configuration file

The scripts in this repository require a database connection string in order to communicate with your database server.

An example config file is provided, named `config.xml.example`. Simply create a copy of this file named "config.xml", and edit it to contain your own database connection string(s).

Some scripts may require additional connection strings, as they may need to communicate with different databases or multiple databases.

If you require assistance figuring out what connection string you need, the following third-party sites may help you:
* https://www.connectionstrings.com/
* https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms722656(v=vs.85)


### Security considerations for config.xml

The above config file is a plaintext file, which may contain a plaintext password on your system. You should take steps to secure this file so that unauthorized users cannot read your passwords from it.
