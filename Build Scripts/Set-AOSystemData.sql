--  Copyright FineBuild Team © 2020.  Distributed under Ms-Pl License
USE [master]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF EXISTS (SELECT 1 FROM sys.procedures WHERE name = N'FB_AGSystemData')
  DROP PROCEDURE dbo.FB_AGSystemData;
GO

CREATE PROC [dbo].[FB_AGSystemData]
 @AGName            NVARCHAR(120)      = '%' -- Name of AG for Failover
,@Execute           CHAR(1)            = 'Y' -- Execute commands
,@TargetServer      NVARCHAR(120)      = ''  -- Internal Use Only
,@RemoteCall        CHAR(1)            = 'N' -- Internal Use Only
,@Operation         CHAR(1)            = ''  -- Internal Use Only
,@Message           VARCHAR(256) = '' OUTPUT -- Status Message
AS
-- FB_AGSystemData
--
--  Copyright FineBuild Team © 2020.  Distributed under Ms-Pl License
--
-- This routine copies System Data (Logins, jobs, etc) from the Primary to all Secondary servers
-- This routine is dependant on dbaTools (https://dbatools.io/)
--
-- Syntax: EXEC FB_AGSystemData @AGName='AGName'
--
-- Do not use any of the parameters marked 'Internal Use Only', they are used within the Main Control Process
--
-- This Proc can be run on either the Primary or a Secondary server in the AG
-- The Main Control Process works out which nodes are Primary and Secondary, and performs the relevant commands on them
--
-- Date        Name  Comment
-- 30/05/2019  EdV   Initial code
-- 24/01/2020  EdV   First FineBuild Version
--
BEGIN;
  SET NOCOUNT ON;

  DECLARE
   @AGWork          NVARCHAR(128)
  ,@JobName         NVARCHAR(128)
  ,@PrimaryServer   NVARCHAR(128)
  ,@SecondaryServer NVARCHAR(128)
  ,@ScheduleId      VARCHAR(8)
  ,@SQLParms        NVARCHAR(256) = '@Message VARCHAR(256) OUTPUT'
  ,@SQLText         NVARCHAR(2000) = '';
 
  DECLARE @Parameters TABLE
  (AGName           NVARCHAR(128)
  ,ExecProcess      CHAR(1)
  ,TargetServer     NVARCHAR(128)
  ,CRLF             CHAR(2)
  ,RemoteCall       CHAR(1)
  ,Operation        CHAR(1));

  INSERT INTO @Parameters (AGName,ExecProcess,TargetServer,CRLF,RemoteCall,Operation)
  VALUES (
   @AGName
  ,@Execute
  ,@TargetServer
  ,Char(13) + Char(10)
  ,@RemoteCall
  ,@Operation);
  UPDATE @Parameters SET
   AGName           = REPLACE(REPLACE(AGName, '[',''),']','')
  FROM @Parameters p;

  DROP TABLE IF EXISTS #AGServers;
  SELECT *
  INTO #AGServers
  FROM dbo.FB_GetAGServers((SELECT AGName FROM @Parameters), '');
  IF (SELECT ExecProcess FROM @Parameters) <> 'Y' SELECT * FROM #AGServers;
  
  IF (SELECT RemoteCall FROM @Parameters) <> 'Y' -- Main Control Process
  BEGIN;

    DECLARE AGNames CURSOR FAST_FORWARD FOR
    SELECT
     AGName, PrimaryServer, SecondaryServer
    FROM #AGServers s
    WHERE TargetServer = 'Y'
    ORDER BY s.AGName;

    OPEN AGNames;
    FETCH NEXT FROM AGNames INTO @AGWork, @PrimaryServer, @SecondaryServer;
    WHILE @@FETCH_STATUS = 0  
    BEGIN;
      SELECT          
       @SQLText   = @SQLText + p.CRLF + 'EXECUTE [' + @PrimaryServer + '].master.dbo.FB_AGSystemData @AGName=''' + @AGWork + ''', @TargetServer=''' + @SecondaryServer + ''', @RemoteCall=''Y'', @Operation = ''M'' '
      FROM @Parameters p
      PRINT @SQLText;
      IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;
      FETCH NEXT FROM AGNames INTO @AGWork, @PrimaryServer, @SecondaryServer;
    END;
    CLOSE AGNames;
    DEALLOCATE AGNames;

    SELECT 
     @SQLText       = 'Processing of Availability Groups complete '
    ,@Message       = 'Completed'
    FROM @Parameters p;
    PRINT @SQLText;
    PRINT '';

  END;


  -- Start of Utility Functions called from the Main Control Process

  IF (SELECT Operation FROM @Parameters) = 'M' -- Main Control Loop
  BEGIN; 
 
    SELECT
     @SQLText     = 'Copy System Data for : ' + p.AGName
    FROM @Parameters p;
    SELECT
     @SQLText     = @SQLText + p.CRLF + 'Current Primary Server: '   + a.PrimaryServer
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName
    WHERE a.TargetServer = 'Y';
    SELECT
     @SQLText     = @SQLText + p.CRLF + 'Current Secondary Server: ' + a.SecondaryServer
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName
    WHERE a.TargetServer = 'Y';
    SELECT
     @SQLText     = @SQLText + p.CRLF + REPLICATE('*', 40)
    ,@Message     = 'Completed'
    FROM @Parameters p;
    PRINT @SQLText;
	
    IF @Message = 'Completed'
    BEGIN;          
      SELECT
       @Message   = 'Copy Critical Data'
      ,@SQLText   = '';
      SELECT          
       @SQLText   = @SQLText + p.CRLF + 'EXECUTE [' + a.PrimaryServer + '].master.dbo.FB_AGSystemData @AGName=''' + a.AGName + ''', @TargetServer=''' + a.SecondaryServer + ''', @RemoteCall=''Y'', @Operation = ''C'', @Message = @Message OUTPUT '
      FROM @Parameters p
      JOIN #AGServers a ON a.AGName = p.AGName;
      PRINT @SQLText;
      IF ((SELECT ExecProcess FROM @Parameters) = 'Y') AND (@SQLText <> '') EXECUTE sp_executeSQL @SQLText, @SQLParms, @Message=@Message OUTPUT;
    END;

    IF @Message = 'Completed'
    BEGIN;          
      SELECT
       @Message   = 'Job Schedule'
      ,@SQLText   = '';
      SELECT          
       @SQLText   = @SQLText + p.CRLF + 'EXECUTE [' + a.SecondaryServer + '].master.dbo.FB_AGSystemData @AGName=''' + p.AGName + ''', @TargetServer=''' + a.SecondaryServer + ''', @RemoteCall=''Y'', @Operation = ''S'', @Message = @Message OUTPUT '
      FROM #AGServers a;
      PRINT @SQLText;
      IF ((SELECT ExecProcess FROM @Parameters) = 'Y') AND (@SQLText <> '') EXECUTE sp_executeSQL @SQLText, @SQLParms, @Message=@Message OUTPUT;
    END;

    IF @Message = 'Completed'
    BEGIN;          
      SELECT
       @Message   = 'Job Definitions'
      ,@SQLText   = '';
      SELECT          
       @SQLText   = @SQLText + p.CRLF + 'EXECUTE [' + a.SecondaryServer + '].master.dbo.FB_AGSystemData @AGName=''' + p.AGName + ''', @TargetServer=''' + a.PrimaryServer + ''', @RemoteCall=''Y'', @Operation = ''J'', @Message = @Message OUTPUT '
      FROM #AGServers a;
      PRINT @SQLText;
      IF ((SELECT ExecProcess FROM @Parameters) = 'Y') AND (@SQLText <> '') EXECUTE sp_executeSQL @SQLText, @SQLParms, @Message=@Message OUTPUT;
    END;

    IF @Message = 'Completed'
    BEGIN;          
      SELECT
       @Message   = 'Schedule Exceptions'
      ,@SQLText   = '';
      SELECT          
       @SQLText   = @SQLText + p.CRLF + 'EXECUTE [' + a.SecondaryServer + '].master.dbo.FB_AGSystemData @AGName=''' + p.AGName + ''', @TargetServer=''' + a.SecondaryServer + ''', @RemoteCall=''Y'', @Operation = ''E'', @Message = @Message OUTPUT '
      FROM #AGServers a;
      PRINT @SQLText;
      IF ((SELECT ExecProcess FROM @Parameters) = 'Y') AND (@SQLText <> '') EXECUTE sp_executeSQL @SQLText, @SQLParms, @Message=@Message OUTPUT;
    END;
	
    SELECT
     @SQLText     = REPLICATE('*', 40) +
                    p.CRLF + 'System Data copy for ' + p.AGName + ' complete'
    FROM @Parameters p;
    PRINT @SQLText;

  END;

  -- Start of Utility Functions called from the Main Control Process

  IF (SELECT Operation FROM @Parameters) = 'C' -- Copy Critical Data
  BEGIN; 

    -- Requires Windows Local Admin access on source server to obtain passwords.  Do not give a Job account this access, instead run this manually on ad-hoc basis
    SELECT
     @Message       = 'Starting Credentials'
    ,@SQLText       = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaCredential       -Source "' +  a.PrimaryServer + '" -Destination "' + p.TargetServer + '" -Force'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
--  IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @Message       = 'Starting Logins'
    ,@SQLText       = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaLogin            -Source "' + a.PrimaryServer + '" -Destination "' + p.TargetServer + '" -Force'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    -- Requires Windows Local Admin access on source server to obtain passwords.  Do not give a Job account this access, instead run this manually on ad-hoc basis
    SELECT
     @Message         = 'Starting Linked Server'
    ,@SQLText         = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaLinkedServer     -Source "' +  a.PrimaryServer + '" -Destination "' + p.TargetServer + '" -Force'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
--  IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @Message       = 'Starting Job Category'
    ,@SQLText       = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaAgentJobCategory -Source "' +  a.PrimaryServer + '" -Destination "' + p.TargetServer + '" -Force'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @Message       = 'Starting Job Operator'
    ,@SQLText       = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaAgentOperator    -Source "' + a.PrimaryServer + '" -Destination "' + p.TargetServer + '" -Force'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @Message       = 'Starting Job Alerts'
    ,@SQLText       = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaAgentAlert       -Source "' + a.PrimaryServer + '" -Destination "' + p.TargetServer + '" -Force'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @Message       = 'Starting Job Proxies'
    ,@SQLText       = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaAgentProxy       -Source "' + a.PrimaryServer + '" -Destination "' + p.TargetServer + '" -Force'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @Message       = 'Starting Job Schedule'
    ,@SQLText       = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaAgentSchedule    -Source "' + a.PrimaryServer + '" -Destination "' + p.TargetServer + '" -Force'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @Message       = 'Starting Mail Profile'
    ,@SQLText      = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaDbMail           -Source "' + a.PrimaryServer + '" -Destination "' + p.TargetServer + '" -Force'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @Message       = 'Starting sp_configure options'
    ,@SQLText       = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaSpConfigure      -Source "' + a.PrimaryServer + '" -Destination "' + p.TargetServer + '"'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @Message       = 'Starting System DB User Objects'
    ,@SQLText       = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaSysDbUserObject  -Source "' + a.PrimaryServer + '" -Destination "' + p.TargetServer + '" -Force'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @Message       = 'Starting Job Definitions'
    ,@SQLText       = 'xp_cmdshell ''MODE CON COLS=120 && POWERSHELL Copy-DbaAgentJob         -Source "' + a.PrimaryServer + '" -Destination "' + p.TargetServer + '" -Force -DisableOnDestination'''
    FROM @Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @SQLText       = 'Copy of Critical Data to server ' + p.TargetServer + ' complete'
    ,@Message       = 'Completed'
    FROM @Parameters p;
    PRINT @SQLText;
    PRINT '';

  END;

  IF (SELECT Operation FROM @Parameters) = 'S' -- Update Schedule data on Target System
  BEGIN; 
    
    DECLARE Job_Schedules CURSOR FAST_FORWARD FOR
    SELECT
     CAST(s.schedule_id AS varchar(8))
    FROM msdb.dbo.sysschedules s
    ORDER BY s.schedule_id;

    OPEN Job_Schedules;
    FETCH NEXT FROM Job_Schedules INTO @ScheduleId;
    WHILE @@FETCH_STATUS = 0  
    BEGIN;
      SELECT 
       @SQLText     = 'EXECUTE msdb.dbo.sp_update_schedule @schedule_id=''' + @Scheduleid + ''',@enabled=0'
      FROM @Parameters p;
      PRINT @SQLText;
      IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;
      FETCH NEXT FROM Job_Schedules INTO @ScheduleId;
    END;
    CLOSE Job_Schedules;
    DEALLOCATE Job_Schedules;

    SELECT 
     @SQLText       = 'Update of Schedule Data on server ' + p.TargetServer + ' complete '
    ,@Message       = 'Completed'
    FROM @Parameters p;
    PRINT @SQLText;
    PRINT '';

  END;

  IF (SELECT Operation FROM @Parameters) = 'J' -- Update Job data on Target System
  BEGIN; 
  
    DECLARE Job_Names CURSOR FAST_FORWARD FOR
    SELECT
     j.name
    FROM msdb.dbo.sysjobs j
    WHERE enabled = 1
    ORDER BY j.name;

    OPEN Job_Names;
    FETCH NEXT FROM Job_Names INTO @JobName;
    WHILE @@FETCH_STATUS = 0  
    BEGIN;
      SELECT 
       @SQLText     = 'IF EXISTS(SELECT 1 FROM [' + p.TargetServer + '].msdb.dbo.sysjobs WHERE name = ''' + @JobName + ''' AND enabled = 1) EXECUTE msdb.dbo.sp_update_job @job_name=''' + @JobName + ''',@enabled=1;'
      FROM @Parameters p;
      PRINT @SQLText;
      IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;
      FETCH NEXT FROM Job_Names INTO @JobName;
    END;
    CLOSE Job_Names;
    DEALLOCATE Job_Names;

    SELECT 
     @SQLText       = 'Update of Job Data on server ' + p.TargetServer + ' complete'
    ,@Message       = 'Completed'
    FROM @Parameters p;
    PRINT @SQLText;
    PRINT '';

  END;

  IF (SELECT Operation FROM @Parameters) = 'E' -- Enable Schedule Exceptions
  BEGIN; 

    SET @SQLText    = '';
    SELECT
     @SQLText       = p.CRLF + @SQLText + 'EXECUTE msdb.dbo.sp_update_schedule @schedule_id=' + Cast(s.schedule_id AS Varchar(8)) + ',@enabled=1; /* ' + j.name + ' */'
    FROM msdb.dbo.sysschedules s
    JOIN msdb.dbo.sysjobschedules js ON js.schedule_id = s.schedule_id
    JOIN msdb.dbo.sysjobs j ON j.job_id = js.job_id
    JOIN master.dbo.FB_AGSystemDataJobExceptions e ON e.JobName = j.name
    CROSS JOIN @Parameters p
    ORDER BY j.name,s.schedule_id;
    PRINT @SQLText;
    IF ((SELECT ExecProcess FROM @Parameters) = 'Y') AND (@SQLText <> '') EXECUTE sp_executeSQL @SQLText;

    SELECT 
     @SQLText       = 'Enable Schedule Exceptions on server ' + p.TargetServer + ' complete'
    ,@Message       = 'Completed'
    FROM @Parameters p;
    PRINT @SQLText;
    PRINT '';

  END;

  RETURN(@@ERROR);

END;

-- Create table for System Data Copy Job Exceptions

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FB_AGSystemDataJobExceptions]') AND type in (N'U'))
  DROP TABLE [dbo].[FB_AGSystemDataJobExceptions];

CREATE TABLE [dbo].[FB_AGSystemDataJobExceptions]
([Id]          INTEGER IDENTITY(1,1)
,[AGName]      NVARCHAR(120) NOT NULL
,[JobName]     NVARCHAR(120) NOT NULL
,CONSTRAINT [PK_AGSystemDataJobExceptions] PRIMARY KEY CLUSTERED ([Id] ASC));

-- Create database roles

EXEC sp_addrole 'FB_AGSystemData',dbo;
GRANT EXECUTE ON dbo.FB_AGSystemData              TO FB_AGSystemData;
GRANT SELECT  ON dbo.FB_AGSystemDataJobExceptions TO FB_AGSystemData;
IF NOT EXISTS (SELECT 1 FROM sys.sysusers WHERE name = '$(strAccount)') EXEC sp_grantdbaccess '$(strAccount)';
EXEC sp_addrolemember 'FB_AGSystemData','$(strAccount)';
GO

USE msdb
GO

EXEC sp_addrole 'FB_AGSystemData',dbo;
GRANT EXECUTE ON dbo.sp_update_job      TO FB_AGSystemData;
GRANT EXECUTE ON dbo.sp_update_schedule TO FB_AGSystemData;
GRANT SELECT  ON dbo.sysjobs            TO FB_AGSystemData;
GRANT SELECT  ON dbo.sysjobschedules    TO FB_AGSystemData;
GRANT SELECT  ON dbo.sysschedules       TO FB_AGSystemData;
IF NOT EXISTS (SELECT 1 FROM sys.sysusers WHERE name = '$(strAccount)') EXEC sp_grantdbaccess '$(strAccount)';
EXEC sp_addrolemember 'FB_AGSystemData','$(strAccount)';
GO

-- Create Job to run System Data Copy

BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0

IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'DBA Tasks' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'DBA Tasks'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'DBA: Copy System Data', 
		@enabled=1, 
		@notify_level_eventlog=0, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'No description available.', 
		@category_name=N'DBA Tasks', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** Object:  Step [Copy System Data]    Script Date: 30/01/2020 14:57:51 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_name=N'DBA: Copy System Data', @step_name=N'Copy System Data', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'EXEC FB_AGSystemData', 
		@database_name=N'master', 
		@flags=4
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_name=N'DBA: Copy System Data', @name=N'Every 3 Hours', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=8, 
		@freq_subday_interval=3, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20000101, 
		@active_end_date=99991231, 
		@active_start_time=0, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:
GO
