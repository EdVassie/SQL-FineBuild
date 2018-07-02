SQL FineBuild ReadMe file
Copyright FineBuild Team © 2018.  Distributed under Ms-Pl License

This file explains the contents of the \FineBuild folder

For more information please see the SQL FineBuild Wiki: https://github.com/SQL-FineBuild/Common/wiki

All of the Example files need the following parameters to be supplied at run time:
/Edition:	The SQL Server Edition (Standard, Enterprise, etc) you want to install
/AdminPassword:	Your own password.  This allows SQL FineBuild tyo log on automatically after a reboot

FineBuild folder contents:

Name				Description
\Build Scripts			Routines used by the SQL FineBuild install process
CmdHere.bat			Command prompt for running files from the FineBuild folder
SQL....Config.xml		FineBuild XML configuration file for a specific version of SQL Server
SQLFineAlwaysOn.bat		Example of installing SQL Serve with Always On (SQL2012 and above only)
SQLFineBuild.bat		The SQL FineBuild install process
SQLFineClient.bat		Example of installing an Administration Server Role build
SQLFineCluster.bat		Example of installing an A/P SQL Server Cluster instance
SQLFineClusterInstance.bat	Example of adding a named instance SQL Cluster to an existing SQL Server Cluster build
SQLFineExpress.bat		Example of installing a SQL Server Express Workstation build
SQLFineFix.bat			Example of installin a SP and/or a CU to an existing SQL Server instance
SQLFineInstance.bat		Example of adding a named instance to an existing SQL Server build
SQLFineServer.bat		Example of installing a complex SQL Server build on to a server
SQLFineWorkstation.bat		Example of installing SQL Servfer to a Workstation