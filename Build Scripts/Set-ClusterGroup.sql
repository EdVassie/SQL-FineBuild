USE [master]
GO

/****** Object:  StoredProcedure [dbo].[FB_ClusterCore]    Script Date: 22/01/2020 12:39:30 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE PROC [dbo].[FB_ClusterCore]
 @Execute           CHAR(1)            = 'Y' -- Execute commands
,@TargetServer      NVARCHAR(120)      = ''  -- Internal Use Only
,@RemoteCall        VARCHAR(1)         = 'N' -- Internal Use Only
,@Operation         CHAR(1)            = ''  -- Internal Use Only
AS
-- FB_ClusterCore
--
-- This routine moves the core Cluster Group to a Secondary node
--
-- Syntax: EXEC FB_ClusterCore
--
-- Do not use any of the parameters marked 'Internal Use Only', they are used within the Main Control Process
--
-- This Proc can be run on either the Primary or a Secondary node in a cluster
-- The Main Control Process works out which nodes are Primary and Secondary, and performs the relevant commands on them
--
-- Date        Name  Comment
-- 04/12/2019  EdV   Initial code
--
BEGIN;
  SET NOCOUNT ON;

  DECLARE
   @JobName         NVARCHAR(128)
  ,@ScheduleId      VARCHAR(8)
  ,@SQLText         NVARCHAR(2000) = '';

  DECLARE @Parameters TABLE
  (AGName           NVARCHAR(128)
  ,ExecProcess      CHAR(1)
  ,PrimaryServer    NVARCHAR(128)
  ,TargetServer     NVARCHAR(128)
  ,CRLF             CHAR(2)
  ,RemoteCall       CHAR(1)
  ,Operation        CHAR(1));

  INSERT INTO @Parameters (ExecProcess,PrimaryServer,TargetServer,CRLF,RemoteCall,Operation)
  VALUES (
   @Execute
  ,@@ServerName
  ,@TargetServer
  ,Char(13) + Char(10)
  ,@RemoteCall
  ,@Operation);
  
  IF (SELECT RemoteCall FROM @Parameters) <> 'Y' -- Main Control Process
  BEGIN;

    SELECT
     @SQLText       = 'Move Core Cluster Resources to secondary server for : ' + @@SERVERNAME
    FROM @Parameters p;
    SELECT
     @SQLText       = @SQLText + p.CRLF + 'Current Primary Server: '   + n.NodeName
    FROM @Parameters p
    CROSS JOIN sys.dm_os_cluster_nodes n
    WHERE n.status = 0 AND n.is_current_owner = 1;
    SELECT TOP 1
     @SQLText       = @SQLText + p.CRLF + 'Target Server         : ' + n.NodeName  
    FROM @Parameters p
    CROSS JOIN sys.dm_os_cluster_nodes n
    WHERE n.status = 0 ORDER BY n.is_current_owner;
    SELECT
     @SQLText       = @SQLText + p.CRLF + REPLICATE('*', 40)
    FROM @Parameters p;
    PRINT @SQLText;

    SET @SQLText    = '';
    SELECT TOP 1    -- Move Cluster Core
     @SQLText       = p.CRLF + 'EXECUTE master.dbo.FB_ClusterCore @TargetServer=''' + n.NodeName + ''', @RemoteCall=''Y'', @Operation = ''C'' '
    FROM @Parameters p
    CROSS JOIN sys.dm_os_cluster_nodes n WHERE n.status = 0 ORDER BY n.is_current_owner;
    PRINT @SQLText;
    IF ((SELECT ExecProcess FROM @Parameters) = 'Y') AND (@SQLText <> '') EXECUTE sp_executeSQL @SQLText;
	
    SELECT
     @SQLText       = REPLICATE('*', 40) +
                      p.CRLF + 'Move Core Cluster Resources for ' + @@Servername + ' complete'
    FROM @Parameters p;
    PRINT @SQLText;

  END;

  -- Start of Utility Functions called from the Main Control Process

  IF (SELECT Operation FROM @Parameters) IN ('C') -- Move Cluster Core
  BEGIN; 

    SELECT @SQLText = 'xp_cmdshell ''CLUSTER GROUP "Cluster Group" /MOVETO:"' + p.TargetServer + '"'''
    FROM @Parameters p;
    PRINT @SQLText;
    IF ((SELECT ExecProcess FROM @Parameters) = 'Y') AND (@SQLText <> '') EXECUTE sp_executeSQL @SQLText;
  
    SELECT @SQLText = 'Move Cluster Quorum on server ' + p.TargetServer
    FROM @Parameters p;
    PRINT @SQLText;
    PRINT '';

  END;

END;
GO

USE [msdb]
GO

/****** Object:  Job [DBA: Move Cluster Core]    Script Date: 22/01/2020 12:45:02 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** Object:  JobCategory [DBA Tasks]    Script Date: 22/01/2020 12:45:02 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'DBA Tasks' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'DBA Tasks'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'DBA: Move Cluster Core', 
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
/****** Object:  Step [Move Cluster Core]    Script Date: 22/01/2020 12:45:02 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'Move Cluster Core', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'EXEC FB_ClusterCore', 
		@database_name=N'master', 
		@flags=4
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'Every 3 Hours', 
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



