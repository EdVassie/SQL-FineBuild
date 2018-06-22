@ECHO OFF
REM Copyright FineBuild Team © 2008 - 2017.  Distributed under Ms-Pl License
REM
CALL "SQLFineBuild.bat" %*     /Type:Full                    /IAcceptLicenseTerms          ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01"   ^
 /Instance:HR                  /TCPPort:1433 ^
 /SetupSQLDBCluster:YES                      ^
 /SetupSQLDBFS:NO                            ^
 /SetupSQLAS:NO                              ^
 /VolProg:C   /VolTempWin:C    /VolDTC:N     ^
 /VolBackup:H /VolData:H /VolDataFT:H /VolLog:I /VolLogTemp:I /VolSysDB:H /VolTemp:H       ^
 /SQLSVCAccount:"ROOT\ServGB_SQLDB_0002"  /SQLSVCPassword:"Kedtt47$jt!rw96vbhH=#tdDrfd4lz" ^
 /AGTSVCACCOUNT:"ROOT\ServGB_SQLAG_0002"  /AGTSVCPASSWORD:"Pff04cnedO#fed$drFExik31da*fgo" ^
 /FTSVCACCOUNT:"ROOT\ServGB_SQLFT_0002"   /FTSVCPASSWORD:"Nyv35$fvsqtvHYw3Seg*$Dpklqr2g9"  ^
 /SETUPCMDSHELL:YES                          ^
 /CMDSHELLACCOUNT:"ROOT\APPGB_SQLCS_0002" /CMDSHELLPASSWORD:"He$dW2zdlh7Ge2cDu0*t" 
REM The following parameters must be suppplied for Windows 2003 Cluster install
REM For details see http://sqlserverfinebuild.codeplex.com/wikipage?title=SQL%20Server%20Cluster%20Install
REM /SQLDOMAINGROUP:"ROOT\GBGGSQLC01DB"         ^
REM /AGTDOMAINGROUP:"ROOT\GBGGSQLC01AGT"