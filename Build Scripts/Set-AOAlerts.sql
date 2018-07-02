--  Copyright FineBuild Team © 2018.  Distributed under Ms-Pl License
USE master
GO

-- Process AG Failover Procedure
IF EXISTS (SELECT 1 FROM sys.procedures WHERE name = N'FB_AGFailover')
  DROP PROCEDURE dbo.FB_AGFailover;
GO

CREATE PROC [dbo].[FB_AGFailover]
AS
--  Copyright FineBuild Team © 2018.  Distributed under Ms-Pl License
BEGIN;
  SET NOCOUNT ON;

  DECLARE 
   @AGName          NVARCHAR(400)
  ,@AlertName       NVARCHAR(400)
  ,@DBName          NVARCHAR(400)
  ,@Role            NVARCHAR(400)
  ,@Enabled         INT
  ,@ScheduleId      VARCHAR(8)
  ,@ServerName      NVARCHAR(400)
  ,@SQLText         NVARCHAR(4000);

  WAITFOR DELAY '00:01'

  SELECT
   @ServerName     = UPPER(CAST(ServerProperty('ServerName') AS NVARCHAR(400)));

  -- Update Log Usage Alerts for Primary/Secondary Server
  DECLARE Log_Alerts CURSOR FAST_FORWARD FOR
  SELECT
   a.name
  ,a.enabled
  ,UPPER(rs.role_desc) AS role_desc
  FROM msdb.dbo.sysalerts a
  CROSS JOIN sys.dm_hadr_availability_replica_states rs
  LEFT JOIN sys.availability_replicas ar
    ON rs.replica_id = ar.replica_id AND rs.group_id = ar.group_id
  WHERE UPPER(ar.replica_server_name) = @ServerName
  AND a.name LIKE 'DB %: Log Usage%'
  ORDER BY a.name;

  OPEN Log_Alerts;
  FETCH NEXT FROM Log_Alerts INTO @AlertName, @Enabled, @Role;
  PRINT 'Server: ' + @ServerName + ' role: ' + @Role;
  WHILE @@FETCH_STATUS = 0  
  BEGIN;
    SELECT 
     @SQLText = 'EXEC msdb.dbo.sp_update_alert @name=''' + @AlertName + ''''
    ,@SQLText = @SQLText + CASE WHEN @Role = 'PRIMARY'   THEN ',@enabled=1;'
                                WHEN @Role = 'SECONDARY' THEN ',@enabled=0;'
                                ELSE ',@enabled=' + Cast(@Enabled AS NVarchar(2)) + ';' END;
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
    FETCH NEXT FROM Log_Alerts INTO @AlertName, @Enabled, @Role;
  END;
  CLOSE Log_Alerts;
  DEALLOCATE Log_Alerts;

  -- Update Job Schedules for Primary/Secondary Server
  DECLARE Job_Schedules CURSOR FAST_FORWARD FOR
  SELECT
   CAST(s.schedule_id AS varchar(8))
  ,s.enabled
  ,UPPER(rs.role_desc) AS role_desc
  FROM msdb.dbo.sysschedules s
  CROSS JOIN sys.dm_hadr_availability_replica_states rs
  LEFT JOIN sys.availability_replicas ar
    ON rs.replica_id = ar.replica_id AND rs.group_id = ar.group_id
  WHERE UPPER(ar.replica_server_name) = @ServerName
  ORDER BY s.schedule_id;

  OPEN Job_Schedules;
  FETCH NEXT FROM Job_Schedules INTO @ScheduleId, @Enabled, @Role;
  PRINT 'Server: ' + @ServerName + ' role: ' + @Role;
  WHILE @@FETCH_STATUS = 0  
  BEGIN;
    SELECT 
     @SQLText = 'EXEC msdb.dbo.sp_update_schedule @schedule_id=' + Cast(@ScheduleId AS NVarchar(8))
    ,@SQLText = @SQLText + CASE WHEN @Role = 'PRIMARY'   THEN ',@enabled=1;'
                                WHEN @Role = 'SECONDARY' THEN ',@enabled=0;'
	                            ELSE ',@enabled=' + Cast(@Enabled AS NVarchar(2)) + ';' END;
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
    FETCH NEXT FROM Job_Schedules INTO @ScheduleId, @Enabled, @Role;
  END;
  CLOSE Job_Schedules;
  DEALLOCATE Job_Schedules;

  -- Update AG Synchronisation for Primary/Secondary Server
  DECLARE AG_DBNames CURSOR FAST_FORWARD FOR
  SELECT
   ag.name 
  ,db.name AS database_name
  FROM master.sys.availability_groups_cluster ag
  JOIN master.sys.dm_hadr_availability_replica_states rs
    ON rs.group_id = ag.group_id AND rs.role_desc = 'SECONDARY'
  JOIN sys.availability_replicas ar
    ON rs.replica_id = ar.replica_id AND rs.group_id = ar.group_id
  JOIN master.sys.availability_databases_cluster dbc
    ON dbc.group_id = ag.group_id
  JOIN master.sys.databases db
    ON db.group_database_id = dbc.group_database_id
  WHERE UPPER(ar.replica_server_name) = @ServerName
  ORDER BY ag.name, db.name;

  OPEN AG_DBNames;
  FETCH NEXT FROM AG_DBNames INTO @AGName, @DBName;
  WHILE @@FETCH_STATUS = 0  
  BEGIN;
    SELECT @SQLText = 'ALTER DATABASE [' + @DBName + '] SET HADR RESUME;';
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
    FETCH NEXT FROM AG_DBNames INTO @AGName, @DBName;
  END;
  CLOSE AG_DBNames;
  DEALLOCATE AG_DBNames;

  -- Update DQ Authorisation for new Primary Server
  IF EXISTS (SELECT 1 FROM master.sys.databases WHERE name = 'DQS_MAIN' AND @Role = 'PRIMARY')
  BEGIN;
    ALTER AUTHORIZATION ON DATABASE::[DQS_MAIN] TO [##MS_dqs_db_owner_login##];
	SELECT @SQLText = 'USE DQS_MAIN;ALTER USER dqs_service WITH LOGIN=[##MS_dqs_service_login##];';
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
  END;

  IF EXISTS (SELECT 1 FROM master.sys.databases WHERE name = 'DQS_PROJECTS' AND @Role = 'PRIMARY')
  BEGIN;
    ALTER AUTHORIZATION ON DATABASE::[DQS_PROJECTS] TO [##MS_dqs_db_owner_login##];
	SELECT @SQLText = 'USE DQS_PROJECTS;ALTER USER dqs_service WITH LOGIN=[##MS_dqs_service_login##];';
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
  END;

  IF EXISTS (SELECT 1 FROM master.sys.databases WHERE name = 'DQS_STAGING_DATA' AND @Role = 'PRIMARY')
  BEGIN;
    ALTER AUTHORIZATION ON DATABASE::[DQS_STAGING_DATA] TO [##MS_dqs_db_owner_login##];
  END;

END;
GO

-- Process Job to run FB_AGFailover
DECLARE 
  @jobId BINARY(16);

IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'DBA Tasks' AND category_class=1)
  EXEC msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'DBA Tasks';

IF EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = N'DBA: AG State Change')
  EXEC msdb.dbo.sp_delete_job @job_name=N'DBA: AG State Change', @delete_unused_schedule=1;

EXEC msdb.dbo.sp_add_job
 @job_name=N'DBA: AG State Change'
,@enabled=1
,@notify_level_eventlog=0
,@notify_level_email=0
,@notify_level_netsend=0
,@notify_level_page=0
,@delete_level=0
,@description=N'Perform required AG Maintenance when AG state changes'
,@category_name=N'DBA Tasks'
,@owner_login_name=N'sa'
,@job_id = @jobId OUTPUT;

EXEC msdb.dbo.sp_add_jobstep
 @job_id=@jobId
,@step_name=N'AG Failover'
,@step_id=1
,@cmdexec_success_code=0
,@on_success_action=1
,@on_success_step_id=0
,@on_fail_action=2
,@on_fail_step_id=0
,@retry_attempts=0
,@retry_interval=0
,@os_run_priority=0
,@subsystem=N'TSQL'
,@command=N'EXEC FB_AGFailover'
,@database_name=N'master'
,@flags=4;

EXEC msdb.dbo.sp_update_job
 @job_id = @jobId
,@start_step_id = 1;

EXEC msdb.dbo.sp_add_jobschedule
 @job_id=@jobId
,@name=N'AG State Change - Startup Check'
,@enabled=1
,@freq_type=64
,@freq_interval=0
,@freq_subday_type=0
,@freq_subday_interval=0
,@freq_relative_interval=0
,@freq_recurrence_factor=0
,@active_start_date=20000101
,@active_end_date=99991231
,@active_start_time=0
,@active_end_time=235959;

EXEC msdb.dbo.sp_add_jobserver
 @job_id = @jobId
,@server_name = N'(local)';

-- Process Alerts to trigger 'DBA: AG State Change' Job

IF EXISTS (SELECT 1 FROM msdb.dbo.sysalerts WHERE name=N'Event - AG Failover')
  EXEC msdb.dbo.sp_delete_alert @name=N'Event - AG Failover';

EXEC msdb.dbo.sp_add_alert @name=N'Event - AG Failover', 
		@message_id=1480, 
		@severity=0, 
		@enabled=1, 
		@delay_between_responses=60, 
		@include_event_description_in=0, 
		@category_name=N'[Uncategorized]', 
		@job_name=N'DBA: AG State Change';

IF EXISTS (SELECT 1 FROM msdb.dbo.sysalerts WHERE name=N'AG State Change')
  EXEC msdb.dbo.sp_delete_alert @name=N'Event - AG State Change';

EXEC msdb.dbo.sp_add_alert @name=N'Event - AG State Change', 
		@message_id=35264, 
		@severity=0, 
		@enabled=1, 
		@delay_between_responses=60, 
		@include_event_description_in=0, 
		@category_name=N'[Uncategorized]', 
		@job_name=N'DBA: AG State Change';


