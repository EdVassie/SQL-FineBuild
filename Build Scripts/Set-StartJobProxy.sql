USE msdb;
GO
CREATE CERTIFICATE StartJobProxy 
  ENCRYPTION BY PASSWORD = 'strsaPwd''
  WITH
   SUBJECT = 'StartJobProxy'
  ,START_DATE='2000/01/01'
  ,EXPIRY_DATE='2999/12/31';
GO
BACKUP CERTIFICATE [StartJobProxy] 
 TO FILE='\\share\SystemDataBackup\server\StartJobProxy.cer'
 WITH PRIVATE KEY (
   FILE = '\\share\SystemDataBackup\server\StartJobProxy.pvk',
   ENCRYPTION BY PASSWORD = 'strsaPwd',
   DECRYPTION BY PASSWORD = 'strsaPwd');
GO
CREATE USER [StartJobProxy] 
 FROM CERTIFICATE [StartJobProxy];
GO
GRANT AUTHENTICATE TO [StartJobProxy];
GRANT EXECUTE ON msdb.dbo.sp_start_job TO [StartJobProxy];
GO


USE [UserDB];
GO
CREATE CERTIFICATE [StartJobProxy] 
 FROM FILE='\\share\SystemDataBackup\server\StartJobProxy.cer'
 WITH PRIVATE KEY (
   FILE = '\\share\SystemDataBackup\server\StartJobProxy.pvk',
   ENCRYPTION BY PASSWORD = 'strsaPwd',
   DECRYPTION BY PASSWORD = 'strsaPwd');
GO
ADD SIGNATURE TO dbo.[spJobStartProc]
  BY CERTIFICATE [StartJobProxy]
  WITH PASSWORD = 'strsaPwd';
GO
EXEC dbo.spJobStartProc;
GO
-- spJobStartProc must be created using WITH EXECUTE AS [Job Owner]