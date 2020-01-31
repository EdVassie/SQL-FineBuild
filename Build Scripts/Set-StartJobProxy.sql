-- Copyright FineBuild Team © 2016 - 2020.  Distributed under Ms-Pl License
-- Process to create a secure Proxy to allow a user to start a SQL Agent job
USE msdb
GO
CREATE CERTIFICATE StartJobProxy 
  ENCRYPTION BY PASSWORD = '$(strStartJobPassword)'
  WITH
   SUBJECT = 'StartJobProxy'
  ,START_DATE='2000/01/01'
  ,EXPIRY_DATE='2999/12/31';
GO
BACKUP CERTIFICATE [StartJobProxy] 
 TO FILE='$(strDirSystemDataBackup)StartJobProxy.cer'
 WITH PRIVATE KEY (
   FILE = '$(strDirSystemDataBackup)StartJobProxy.pvk',
   ENCRYPTION BY PASSWORD = '$(strStartJobPassword)',
   DECRYPTION BY PASSWORD = '$(strStartJobPassword)');
GO
CREATE USER [StartJobProxy] 
 FROM CERTIFICATE [StartJobProxy];
GO
GRANT AUTHENTICATE TO [StartJobProxy];
GRANT EXECUTE ON msdb.dbo.sp_start_job TO [StartJobProxy];
GO

/*
USE [UserDB]
GO
CREATE CERTIFICATE [StartJobProxy] 
 FROM FILE='$(strDirSystemDataBackup)StartJobProxy.cer'
 WITH PRIVATE KEY (
   FILE = '$(strDirSystemDataBackup)StartJobProxy.pvk',
   ENCRYPTION BY PASSWORD = '$(strStartJobPassword)',
   DECRYPTION BY PASSWORD = '$(strStartJobPassword)');
GO
*/

/*
-- Example procedure that uses the StartJobProxy
-- [Job Owner] must be the same account that owns 'User Job'
USE [UserDB]
GO
CREATE PROCEDURE dbo.[spUserDBStartJobProc]
WITH EXECUTE AS [Job Owner]
AS
BEGIN;
  EXEC msdb.dbo.sp_start_job @job_name='User Job';
END;
GO
ADD SIGNATURE TO dbo.[spUserDBStartJobProc]
  BY CERTIFICATE [StartJobProxy]
  WITH PASSWORD = '$(strStartJobPassword)';
GO
-- GRANT EXECUTE ON dbo.[spUserDBStartJobProc] TO [Required User or Role];
--EXEC dbo.spUserDBStartJobProc;
*/