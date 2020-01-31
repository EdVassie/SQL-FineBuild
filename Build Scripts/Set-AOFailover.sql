--  Copyright FineBuild Team © 2018-2020.  Distributed under Ms-Pl License
USE master
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- Process FB_GetAGServers Function
IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_TYPE = 'FUNCTION' AND ROUTINE_NAME = N'FB_GetAGServers')
  DROP FUNCTION dbo.FB_GetAGServers;
GO

CREATE FUNCTION [dbo].[FB_GetAGServers] 
   (@AGName         NVARCHAR(128)
   ,@TargetServer   NVARCHAR(128))

RETURNS @AGServers  TABLE
  (AGName           NVARCHAR(128)
  ,AGType           CHAR(1)
  ,AvailabilityMode INT
  ,RequiredCommit   INT
  ,PrimaryServer    NVARCHAR(128)
  ,SecondaryServer  NVARCHAR(128)
  ,Endpoint         NVARCHAR(128)
  ,ServerId         INT
  ,TargetServer     CHAR(1)) 

AS
-- FB_GetAGServers
--
--  Copyright FineBuild Team © 2019.  Distributed under Ms-Pl License
--
-- Get list of Servers and their roles in an Availability Group
-- The routine will work out from the AG Name what type of availability group is involved, and process accordingly
--
-- Syntax: SELECT dbo.FB_GetAGServers('%','')      to automatically select first Secondary as the TargetServer for each Availability Group
--     or: SELECT dbo.FB_GetAGServers('AGName','') to automatically select first Secondary as the TargetServer for given Availability Group
--     or: SELECT dbo.FB_GetAGServers('AGName','TargetServer') if a specific target server is required for given Availability Group
--
-- Date        Name  Comment
-- 12/11/2019  EdV   Initial code
-- 28/01/2020  EdV   Initial FineBuild Version
--
BEGIN;

  DECLARE @Parameters TABLE
  (AGName           NVARCHAR(128)
  ,TargetServer     NVARCHAR(128));

  INSERT INTO @Parameters (AGName,TargetServer) 
  VALUES(@AGName,@TargetServer);

  INSERT INTO @AGServers (AGName, AGType, AvailabilityMode, RequiredCommit, PrimaryServer, SecondaryServer, Endpoint, ServerId)
  SELECT
   ag.name
  ,CASE WHEN ag.is_distributed = 1 THEN 'D' WHEN ag.basic_features = 1 THEN 'B' WHEN ag.cluster_type_desc = 'none' THEN 'N' ELSE 'C' END AS AGType
  ,ars.availability_mode
  ,ISNULL(ag.required_synchronized_secondaries_to_commit, 0)
  ,arp.replica_server_name
  ,ars.replica_server_name
  ,SUBSTRING(ISNULL(ars.endpoint_url, arp.endpoint_url), 7, CHARINDEX('.', ISNULL(ars.endpoint_url, arp.endpoint_url)) - 7)
  ,ROW_NUMBER() OVER(PARTITION BY AGName ORDER BY ars.replica_server_name)
  FROM @Parameters p
  JOIN [sys].[availability_groups] ag ON ag.name LIKE p.AGName
  LEFT JOIN [sys].[availability_replicas] ars ON ars.group_id = ag.group_id
  LEFT JOIN [sys].[dm_hadr_availability_replica_states] arss ON arss.group_id = ag.group_id AND arss.replica_id = ars.replica_id AND arss.role <> 1
  LEFT JOIN [sys].[availability_replicas] arp ON arp.group_id = ag.group_id
  LEFT JOIN [sys].[dm_hadr_availability_replica_states] arps ON arps.group_id = ag.group_id AND arps.replica_id = arp.replica_id AND arps.role = 1;

  UPDATE @AGServers SET
   TargetServer = 'Y'
  FROM @Parameters p
  WHERE ServerId = (SELECT MAX(ServerId) FROM @AGServers WHERE (ServerId = 1) OR (p.TargetServer = Endpoint) GROUP BY AGName);

  RETURN;

END;
GO

-- Process FB_AGFailover Procedure
IF EXISTS (SELECT 1 FROM sys.procedures WHERE name = N'FB_AGFailover')
  DROP PROCEDURE dbo.FB_AGFailover;
GO

CREATE       PROC [dbo].[FB_AGFailover]
 @AGName            NVARCHAR(120)      = '%' -- Name of AG for Failover
,@TargetServer      NVARCHAR(128)      = ''  -- Name of desired New Primary server
,@Execute           CHAR(1)            = 'Y' -- Execute commands
,@RemoteCall        VARCHAR(1)         = 'N' -- Internal Use Only
,@Operation         CHAR(1)            = ''  -- Internal Use Only
AS
-- FB_AGFailover
--
--  Copyright FineBuild Team © 2020.  Distributed under Ms-Pl License
--
-- This routine performs a failover of an Availability Group
-- The routine will work out from the AG Name what type of availability group is involved, and process accordingly
--
-- Syntax: EXEC FB_AGFailover @AGName='???'
--     or: EXEC FB_AGFailover @AGName='???', @TargetServer='???' if a specific target server is required
--
-- Do not use any of the parameters marked 'Internal Use Only', they are used within the Main Control Process
--
-- This Proc can be run on either the Primary or a Secondary server in the AG
-- The Main Control Process works out which nodes are Primary and Secondary, and performs the relevant commands on them
--
-- Date        Name  Comment
-- 30/05/2019  EdV   Initial code
-- 25/06/2019  EdV   Added @TargetServer and @Execute logic
-- 28/01/2020  EdV   Added @TargetServer and @Execute logic
--
BEGIN;
  SET NOCOUNT ON;

  DECLARE
   @Server          NVARCHAR(128)
  ,@SQLText         NVARCHAR(2000) = ''
  ,@NonSync         INT            = -1;

  DROP TABLE IF EXISTS #Parameters;
  CREATE TABLE #Parameters
  (AGName           NVARCHAR(128)
  ,TargetServer     NVARCHAR(128)
  ,ExecProcess      CHAR(1)
  ,CRLF             CHAR(2)
  ,RemoteCall       CHAR(1)
  ,Operation        CHAR(1));

  INSERT INTO #Parameters (AGName,TargetServer,ExecProcess,CRLF,RemoteCall,Operation)
  VALUES (
   @AGName
  ,@TargetServer
  ,@Execute
  ,Char(13) + Char(10)
  ,@RemoteCall
  ,@Operation);
  UPDATE #Parameters SET
   AGName           = REPLACE(REPLACE(AGName, '[',''),']','')
  ,TargetServer     = CASE WHEN p.TargetServer <> '' THEN p.TargetServer ELSE CAST(SERVERPROPERTY('ServerName') AS NVARCHAR(128)) END
  FROM #Parameters p;

  DROP TABLE IF EXISTS #AGServers;
  SELECT *
  INTO #AGServers
  FROM dbo.FB_GetAGServers((SELECT AGName FROM #Parameters), (SELECT TargetServer FROM #Parameters));

  IF (SELECT ExecProcess FROM #Parameters) <> 'Y' SELECT * FROM #AGServers;

  IF (SELECT RemoteCall FROM #Parameters) <> 'Y' -- Main Control Process
  BEGIN;
 
    SELECT
     @SQLText          = 'Performing Failover of: ' + p.AGName
    FROM #Parameters p;
    SELECT
     @SQLText          = @SQLText + p.CRLF + 'Current Primary Server: '   + a.ServerName
    FROM #Parameters p
    JOIN #AGServers a ON a.AGName LIKE p.AGName
    WHERE a.ServerRole = 'P';
    SELECT
     @SQLText          = @SQLText + p.CRLF + 'Current Secondary Server: ' + a.ServerName  
    FROM #Parameters p
    JOIN #AGServers a ON a.AGName LIKE p.AGName
    WHERE a.TargetServer = 'Y';
    SELECT
     @SQLText          = @SQLText + p.CRLF + REPLICATE('*', 40)
    FROM #Parameters p;
    PRINT @SQLText;
    
    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.ServerName + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @RemoteCall=''Y'',@Operation = ''S'''
    FROM #Parameters p
    JOIN #AGServers a ON a.AGName LIKE p.AGName
    WHERE a.ServerRole = 'P';
    IF (SELECT ExecProcess FROM #Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;
    IF (SELECT ExecProcess FROM #Parameters) <> 'Y' PRINT @SQLText;

    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.ServerName + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @RemoteCall=''Y'',@Operation = ''S'''
    FROM #Parameters p
    JOIN #AGServers a ON a.AGName LIKE p.AGName
    WHERE a.ServerRole = 'S';
    IF (SELECT ExecProcess FROM #Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;
    IF (SELECT ExecProcess FROM #Parameters) <> 'Y' PRINT @SQLText;
	
    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.ServerName + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @RemoteCall=''Y'',@Operation = ''R'''
    FROM #Parameters p
    JOIN #AGServers a ON a.AGName LIKE p.AGName
    WHERE a.ServerRole = 'P';
    IF (SELECT ExecProcess FROM #Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;
    IF (SELECT ExecProcess FROM #Parameters) <> 'Y' PRINT @SQLText;
 
    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.ServerName + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @RemoteCall=''Y'',@Operation = ''F'''
    FROM #Parameters p
    JOIN #AGServers a ON a.AGName LIKE p.AGName
    WHERE a.TargetServer = 'Y';
    IF (SELECT ExecProcess FROM #Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;
    IF (SELECT ExecProcess FROM #Parameters) <> 'Y' PRINT @SQLText;

    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.ServerName + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @RemoteCall=''Y'',@Operation = ''A'''
    FROM #Parameters p
    JOIN #AGServers a ON a.AGName LIKE p.AGName
    WHERE a.ServerRole = 'P' AND (a.AvailabilityMode = 0 OR a.AGType = 'D');
    IF (SELECT ExecProcess FROM #Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;
    IF (SELECT ExecProcess FROM #Parameters) <> 'Y' PRINT @SQLText;

    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.ServerName + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @RemoteCall=''Y'',@Operation = ''A'''
    FROM #Parameters p
    JOIN #AGServers a ON a.AGName LIKE p.AGName
    WHERE a.ServerRole = 'S' AND (a.AvailabilityMode = 0 OR a.AGType = 'D');
    IF (SELECT ExecProcess FROM #Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;
    IF (SELECT ExecProcess FROM #Parameters) <> 'Y' PRINT @SQLText;
	
    SELECT
     @SQLText       = REPLICATE('*', 40) +
                      p.CRLF + 'SQL Failover of ' + p.AGName + ' to ' + a.ServerName + ' complete'
    FROM #Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName
    WHERE a.TargetServer = 'Y';
    SELECT
     @SQLText       = @SQLText +
                      p.CRLF + 'Update DNS Alias for ' + p.AGName + ' to point to ' + a.ServerName
    FROM #Parameters p
    JOIN #AGServers a ON a.AGName = p.AGName
    WHERE a.TargetServer = 'Y' AND a.AGType = 'D';
    PRINT @SQLText;

  END;

  -- Start of Utility Functions called from the Main Control Process

  IF (SELECT Operation FROM #Parameters) IN ('S','A') -- AG Communication Mode
  BEGIN; 
    SELECT
     @SQLText       = 'ALTER AVAILABILITY GROUP [' + p.AGName + '] ' +
                      p.CRLF + '  MODIFY AVAILABILITY GROUP ON ' 
    FROM #Parameters p;
    SELECT
     @SQLText       = @SQLText + p.CRLF + '  ''' + ar.replica_server_name + ''' '
    ,@SQLText       = @SQLText + 'WITH (AVAILABILITY_MODE=' + CASE WHEN p.Operation = 'S' THEN 'SYNCHRONOUS' ELSE 'ASYNCHRONOUS' END + '_COMMIT),'
    FROM #Parameters p
    JOIN [master].[sys].[availability_groups] ag ON ag.name = p.AGName
    JOIN [master].[sys].[availability_replicas] ar ON ar.group_id = ag.group_id;
    SELECT @SQLText = LEFT(@SQLText, LEN(@SQLText) - 1) + ';';
    PRINT @SQLText;
    EXECUTE sp_executeSQL @SQLText;
    PRINT 'Processing on Server ' + @Server;
    PRINT '';
  END;

  IF (SELECT Operation FROM #Parameters) = 'R' -- AG Role Change
  BEGIN; 
    WHILE (@NonSync <> 0)
    BEGIN;
      SELECT
       @SQLText     = CASE WHEN @NonSync > 0 THEN 'Waiting for Primary states to be SYNCHRONIZED for ' + Cast(@NonSync AS varchar(8)) + ' databases'
                           ELSE '' END;
      IF @SQLText <> '' PRINT @SQLText;
      WAITFOR DELAY '00:00:01'
      SELECT
       @NonSync = COUNT(*)
      FROM #Parameters p
      JOIN [master].[sys].[availability_groups] ag ON ag.name = p.AGName
      JOIN [master].[sys].[availability_replicas] ar1 ON ar1.group_id = ag.group_id
      JOIN [master].[sys].[dm_hadr_database_replica_states] drs1 ON drs1.group_id = ar1.group_id AND drs1.replica_id = ar1.replica_id 
      JOIN [master].[sys].[availability_replicas] ar2 ON ar2.group_id = ar1.group_id AND ar2.replica_id <> ar1.replica_id
      JOIN [master].[sys].[availability_replicas] ar3 ON ar3.endpoint_url = ar2.endpoint_url
      JOIN [master].[sys].[dm_hadr_database_replica_states] drs2 ON drs2.group_id = ar3.group_id AND drs2.replica_id = ar3.replica_id AND drs2.database_id = drs1.database_id
      WHERE drs2.synchronization_state_desc NOT IN ('SYNCHRONIZED', 'NOT SYNCHRONIZING') OR drs2.last_commit_lsn <> drs1.last_commit_lsn;
    END;

    SELECT
     @SQLText       = 'ALTER AVAILABILITY GROUP [' + p.AGName + '] SET (ROLE=SECONDARY);' 
    FROM #Parameters p;
    PRINT @SQLText;
    EXECUTE sp_executeSQL @SQLText;
    PRINT 'Processing on Server ' + @Server;
    PRINT '';
  END;

  IF (SELECT Operation FROM #Parameters) = 'F' -- AG Failover
  BEGIN; 
    WHILE (@NonSync <> 0)
    BEGIN;
      SELECT
       @SQLText     = @SQLText + CASE WHEN @NonSync > 0 THEN 'Waiting for Secondary states to be SYNCHRONIZED for ' + Cast(@NonSync AS varchar(8)) + ' databases'
                                 ELSE '' END;
      IF @SQLText <> '' PRINT @SQLText;
      WAITFOR DELAY '00:00:01'
      SELECT
       @NonSync = COUNT(*)
      FROM #Parameters p
      JOIN [master].[sys].[availability_groups] ag ON ag.name = p.AGName
      JOIN [master].[sys].[availability_replicas] ar1 ON ag.group_id = ag.group_id
      JOIN [master].[sys].[availability_replicas] ar2 ON ar2.endpoint_url = ar1.endpoint_url
      JOIN [master].[sys].[availability_groups] ag2 ON ag2.group_id = ar2.group_id AND ag2.name = ar1.replica_server_name
      JOIN [master].[sys].[dm_hadr_database_replica_states] drs ON drs.group_id = ar2.group_id
      WHERE drs.synchronization_state_desc NOT IN ('SYNCHRONIZED', 'NOT SYNCHRONIZING');
    END;
    SELECT
     @SQLText       = 'ALTER AVAILABILITY GROUP [' + p.AGName + '] FORCE_FAILOVER_ALLOW_DATA_LOSS;' 
    FROM #Parameters p;
    PRINT @SQLText;
    EXECUTE sp_executeSQL @SQLText;
    PRINT 'Processing on Server ' + @Server;
    PRINT '';
  END;

END;
GO

-- Process FB_AGPostFailover Procedure
IF EXISTS (SELECT 1 FROM sys.procedures WHERE name = N'FB_AGPostFailover')
  DROP PROCEDURE dbo.FB_AGPostFailover;
GO

CREATE PROC [dbo].[FB_AGPostFailover]
AS
-- FB_AGPostFailover
--
--  Copyright FineBuild Team © 2020.  Distributed under Ms-Pl License
--
-- This routine performs post-failover tasks for an Availability Group
--
-- Syntax: EXEC FB_AGPostFailover
--
-- Date        Name  Comment
-- 25/06/2019  EdV   Initial code
-- 28/01/2020  EdV   Initial FineBuild Version
--
BEGIN;
  SET NOCOUNT ON;

  DECLARE 
   @AGName          NVARCHAR(400)
  ,@AlertName       NVARCHAR(400)
  ,@DBName          NVARCHAR(400)
  ,@Owner           NVARCHAR(400)
  ,@Role            NVARCHAR(400)
  ,@Enabled         INT
  ,@ScheduleId      VARCHAR(8)
  ,@ServerName      NVARCHAR(400)
  ,@SQLText         NVARCHAR(4000);

  CREATE TABLE #SpecialDB
   (Name           NVARCHAR(120));

  INSERT INTO #SpecialDB (Name)
  SELECT 'master'
  UNION ALL
  SELECT 'model'
  UNION ALL
  SELECT 'msdb'
  UNION ALL
  SELECT 'tempdb'
  UNION ALL
  SELECT 'mssqlsystemresource'
  UNION ALL
  SELECT 'SSISDB'
  UNION
  SELECT name FROM master.sys.databases WHERE is_distributor = 1;

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
  AND HAS_DBACCESS(SUBSTRING(LEFT(performance_condition, CHARINDEX('|>|', performance_condition) - 1), LEN('SQLServer:Databases|Percent Log Used|') + 1, LEN(performance_condition))) = 1
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
     @SQLText = 'EXEC msdb.dbo.sp_update_schedule @schedule_id=' + @ScheduleId
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
  ,UPPER(rs.role_desc) AS role_desc
  ,CASE WHEN sp.name > '' THEN 'sa' -- EMV 14/05/19
        WHEN db.name LIKE 'DQS_%' THEN '##MS_dqs_db_owner_login##' -- EMV 16/05/19
        WHEN c.name > '' THEN c.credential_identity -- EMV 14/05/19
        ELSE 'sa' END -- EMV 14/05/19 
  FROM master.sys.availability_groups_cluster ag
  JOIN master.sys.dm_hadr_availability_replica_states rs
    ON rs.group_id = ag.group_id
  JOIN sys.availability_replicas ar
    ON rs.replica_id = ar.replica_id AND rs.group_id = ar.group_id
  JOIN master.sys.availability_databases_cluster dbc
    ON dbc.group_id = ag.group_id
  JOIN master.sys.databases db
    ON db.group_database_id = dbc.group_database_id
  LEFT OUTER JOIN sys.credentials c
    ON c.name = 'StandardDBOwner'
  LEFT OUTER JOIN #SpecialDB sp
    ON sp.name = db.name
  WHERE UPPER(ar.replica_server_name) = @ServerName
  ORDER BY ag.name, db.name;

  OPEN AG_DBNames;
  FETCH NEXT FROM AG_DBNames INTO @AGName, @DBName, @Role, @Owner;
  WHILE @@FETCH_STATUS = 0  
  BEGIN;
    SELECT @SQLText = CASE WHEN @Role = 'SECONDARY' THEN 'ALTER DATABASE [' + @DBName + '] SET HADR RESUME;'
	                       ELSE 'ALTER AUTHORIZATION ON DATABASE::[' + @DBName + '] TO [' + @Owner + '];' END;
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
    FETCH NEXT FROM AG_DBNames INTO @AGName, @DBName, @Role, @Owner;
  END;
  CLOSE AG_DBNames;
  DEALLOCATE AG_DBNames;

  -- Update DQ Authorisation for new Primary Server
  IF EXISTS (SELECT 1 FROM master.sys.databases WHERE name = 'DQS_MAIN' AND @Role = 'PRIMARY')
  BEGIN;
	SELECT @SQLText = 'USE DQS_MAIN;ALTER USER dqs_service WITH LOGIN=[##MS_dqs_service_login##];';
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
  END;

  IF EXISTS (SELECT 1 FROM master.sys.databases WHERE name = 'DQS_PROJECTS' AND @Role = 'PRIMARY')
  BEGIN;
	SELECT @SQLText = 'USE DQS_PROJECTS;ALTER USER dqs_service WITH LOGIN=[##MS_dqs_service_login##];';
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
  END;

END;
GO

-- Process Job to run FB_AGPostFailover
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
,@step_name=N'AG Post Failover'
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
,@command=N'EXEC FB_AGPostFailover'
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

GO