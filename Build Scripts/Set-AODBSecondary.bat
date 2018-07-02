@ECHO OFF
REM Restore AO Secondary DB 
REM Copyright FineBuild Team © 2018.  Distributed under Ms-Pl License
REM
REM Created 31 Mar 2018 by Ed Vassie V1.0
REM
SET FBAODBCmdSQL=%1
SET FBAODBServer=%2
SET FBAODBName=%3
SET FBAODBSSIS=%4
SET FBAODBSSISPwd=%5

SET FBAODBCmd="EXEC master.dbo.FB_DBRestore @DatabaseName='%FBAODBName%',@Server='%FBAODBServer%',@Execute='Y'"
SET FBAODBCmdSSIS="EXEC master.dbo.sp_control_dbmasterkey_password @db_name='%FBAODBName%',@password='%FBAODBSSISPwd%',@action='add'"

ECHO %FBAODBCmdSQL% -Q %FBAODBCmd%
%FBAODBCmdSQL% -Q %FBAODBCmd%

IF '%FBAODBName%'=='%FBAODBSSIS%' (
  %FBAODBCmdSQL% -Q %FBAODBCmdSSIS%
)

SET FBAODBCmdSQL=
SET FBAODBCmd=
SET FBAODBCmdSSIS=
SET FBAODBServer=
SET FBAODBName=
SET FBAODBSSIS=
SET FBAODBSSISPwd=