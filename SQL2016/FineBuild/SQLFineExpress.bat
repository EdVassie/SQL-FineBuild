@ECHO OFF
REM Copyright FineBuild Team © 2015 - 2016.  Distributed under Ms-Pl License
REM
CALL "SQLFineBuild.bat" %*     /Type:Workstation             /IAcceptLicenseTerms          ^
 /Edition:Express                       ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"