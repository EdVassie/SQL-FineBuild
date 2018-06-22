@ECHO OFF
REM Copyright FineBuild Team © 2008 - 2017.  Distributed under Ms-Pl License
REM
CALL "SqlFineBuild.bat" %*     /Type:Full                      /IAcceptLicenseTerms        ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"          /GroupDBANonSA:"GBGGDBAN01" ^
 /Instance:HR                  /TCPPort:1433 ^
 /SetupSQLDBCluster:YES                      ^
 /SetupSQLDBFS:NO                            ^
 /SetupSQLAS:NO                              ^
 /VolProg:C   /VolTempWin:C    /VolDTC:N     ^
 /VolBackup:H /VolData:H /VolDataFT:H /VolLog:I /VolLogTemp:I /VolSysDB:H /VolTemp:H       ^
 /SQLAccount:"ROOT\ServGB_SQLDB_0002"        /SQLPassword:"Kedtt47$jt!rw96vbhH=#tdDrfd4lz" ^
 /AGTACCOUNT:"ROOT\ServGB_SQLAG_0002"        /AGTPASSWORD:"Pff04cnedO#fed$drFExik31da*fgo" ^
 /SETUPCMDSHELL:YES                          ^
 /CMDSHELLACCOUNT:"ROOT\APPGB_SQLCS_0002"    /CMDSHELLPASSWORD:"He$dW2zdlh7Ge2cDu0*t"      ^
 /SQLCLUSTERGROUP:"ROOT\GBGGSQLC01DB"        /AGTCLUSTERGROUP:"ROOT\GBGGSQLC01AGT"         ^
 /ASCLUSTERGROUP:"ROOT\GBGGSQLC01AS"         /FTSCLUSTERGROUP:"ROOT\GBGGSQLC01FTS"
REM
REM /AdminPassword: must be supplied and contain the password for the account running SQL FineBuild
REM For details see http://sqlserverfinebuild.codeplex.com/wikipage?title=SQL%20Server%20Cluster%20Install