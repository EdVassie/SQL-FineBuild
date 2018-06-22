@ECHO OFF
REM Copyright FineBuild Team © 2008 - 2016.  Distributed under Ms-Pl License
REM
CALL "SqlFineBuild.bat" %*     /Type:Full                    /IAcceptLicenseTerms             ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01"      ^
 /Instance:HR                  /TCPPort:8950                 ^
 /VolProg:C                    ^
 /VolBackup:I /VolData:JF /VolDataFT:F       /VolLog:KG /VolTemp:T                            ^
 /SQLSVCAccount:"ROOT\ServGB_SQLDB_0002"     /SQLSVCPassword:"Kedtt47$jt!rw96vbhH=#tdDrfd4lz" ^
 /AGTSVCACCOUNT:"ROOT\ServGB_SQLAG_0002"     /AGTSVCPASSWORD:"Pff04cnedO#fed$drFExik31da*fgo" ^
 /FTSVCACCOUNT:"ROOT\ServGB_SQLFT_0002"      /FTSVCPASSWORD:"Nyv35$fvsqtvHYw3Seg*$Dpklqr2g9"  ^
 /SETUPCMDSHELL:YES ^
 /CMDSHELLACCOUNT:"ROOT\APPGB_SQLCS_0002"    /CMDSHELLPASSWORD:"He$dW2zdlh7Ge2cDu0*t"
