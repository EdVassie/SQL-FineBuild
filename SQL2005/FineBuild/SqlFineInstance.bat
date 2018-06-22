@ECHO OFF
REM Copyright FineBuild Team © 2008 - 2016.  Distributed under Ms-Pl License
REM
CALL "SqlFineBuild.bat" %*     /Type:Full                    /IAcceptLicenseTerms        ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01" ^
 /Instance:HR                  /TCPPort:8950                                             ^
 /PIDKEY:"00000-00000-00000-00000-00000"   ^
 /SQLAccount:"ROOT\ServGB_SQLDB_0002"      /SQLPassword:"Kedtt47$jt!rw96vbhH=#tdDrfd4lz" ^
 /AGTACCOUNT:"ROOT\ServGB_SQLAG_0002"      /AGTPASSWORD:"Pff04cnedO#fed$drFExik31da*fgo" ^
 /SETUPCMDSHELL:YES ^
 /CMDSHELLACCOUNT:"ROOT\APPGB_SQLCS_0002"  /CMDSHELLPASSWORD:"He$dW2zdlh7Ge2cDu0*t"
