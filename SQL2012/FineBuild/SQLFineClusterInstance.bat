@ECHO OFF
REM Copyright FineBuild Team © 2012 - 2017.  Distributed under Ms-Pl License
REM
CALL "SQLFineBuild.bat" %*     /Type:Full                    /IAcceptLicenseTerms          ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01"   ^
 /Instance:HR                  /TCPPort:1433 ^
 /SetupSQLDBCluster:YES                      ^
 /SetupSQLDBFS:NO                            ^
 /SetupSQLAS:NO                              ^
 /SetupAlwaysOn:Yes                          ^
 /SQLSVCAccount:"ROOT\ServGB_SQLDB_1$"       ^
 /AGTSVCACCOUNT:"ROOT\ServGB_SQLAG_1$"       ^
 /FTSVCACCOUNT:"ROOT\ServGB_SQLFT_1$"        ^
 /VolProg:C   /VolTempWin:C    /VolDTC:N     ^
 /VolBackup:I /VolData:H /VolLog:I           ^
 /SETUPCMDSHELL:YES                          ^
 /CMDSHELLACCOUNT:"ROOT\APPGB_SQLCS_0002" /CMDSHELLPASSWORD:"He$dW2zdlh7Ge2cDu0*t" 