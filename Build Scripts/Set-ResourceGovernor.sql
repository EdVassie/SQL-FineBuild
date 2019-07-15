-- Copyright FineBuild Team © 2016 - 2018.  Distributed under Ms-Pl License
USE [master]
GO
-- Get variable data
DECLARE
 @MaxDop                        varchar(4)
,@RPOptions                     varchar(100)
,@SQL                           Varchar(max)
,@SQLVersion                    Varchar(8);

SELECT
 @MaxDop			= CAST(value AS varchar(4))
,@RPOptions                     = ''
,@RPOptions                     = CASE WHEN CAST(SERVERPROPERTY('ProductVersion') AS CHAR(2)) >= '11' THEN @RPOptions + ',cap_cpu_percent=100,AFFINITY SCHEDULER = AUTO' ELSE @RPOptions END
,@RPOptions                     = CASE WHEN CAST(SERVERPROPERTY('ProductVersion') AS CHAR(2)) >= '12' THEN @RPOptions + ',min_iops_per_volume=0,max_iops_per_volume=0'   ELSE @RPOptions END
FROM sys.configurations
WHERE name = 'max degree of parallelism';

-- Setup Resource Governor Pools

IF CAST(SERVERPROPERTY('ProductVersion') AS CHAR(2)) >= '13'
  IF NOT EXISTS ( SELECT name FROM sys.resource_governor_external_resource_pools WHERE name = N'Analytics')
  BEGIN;
    SELECT @SQL = 
    'CREATE EXTERNAL RESOURCE POOL [Analytics]
      WITH(max_cpu_percent=100
	,max_memory_percent=100
	,AFFINITY CPU = AUTO
	,max_processes=0)';
  EXECUTE (@SQL);
  END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_resource_pools WHERE name = N'Analytics')
BEGIN;
  SELECT @SQL = 
  'CREATE RESOURCE POOL [Analytics]
     WITH(min_cpu_percent=0
	,max_cpu_percent=100
	,min_memory_percent=0
	,max_memory_percent=100'
	+ @RPOptions + ')';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_resource_pools WHERE name = N'AdHoc')
BEGIN;
  SELECT @SQL = 
  'CREATE RESOURCE POOL [AdHoc]
     WITH(min_cpu_percent=0
	,max_cpu_percent=100
	,min_memory_percent=0
	,max_memory_percent=100'
	+ @RPOptions + ')';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_resource_pools WHERE name = N'Batch')
BEGIN;
  SELECT @SQL = 
  'CREATE RESOURCE POOL [Batch]
     WITH(min_cpu_percent=0
	,max_cpu_percent=100
	,min_memory_percent=0
	,max_memory_percent=100'
	+ @RPOptions + ')';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_resource_pools WHERE name = N'DBA')
BEGIN;
  SELECT @SQL = 
  'CREATE RESOURCE POOL [DBA]
     WITH(min_cpu_percent=0
	,max_cpu_percent=100
	,min_memory_percent=0
	,max_memory_percent=100'
	+ @RPOptions + ')';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_resource_pools WHERE name = N'OLTP')
BEGIN;
  SELECT @SQL = 
  'CREATE RESOURCE POOL [OLTP]
     WITH(min_cpu_percent=0
	,max_cpu_percent=100
	,min_memory_percent=0
	,max_memory_percent=100'
	+ @RPOptions + ')';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_resource_pools WHERE name = N'Report')
BEGIN;
  SELECT @SQL = 
  'CREATE RESOURCE POOL [Report]
     WITH(min_cpu_percent=0
	,max_cpu_percent=100
	,min_memory_percent=0
	,max_memory_percent=100'
	+ @RPOptions + ')';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_resource_pools WHERE name = N'System')
BEGIN;
  SELECT @SQL = 
  'CREATE RESOURCE POOL [System]
     WITH(min_cpu_percent=0
	,max_cpu_percent=100
	,min_memory_percent=0
	,max_memory_percent=100'
	+ @RPOptions + ')';
  EXECUTE (@SQL);
END;

-- Setup Workload Groups

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_workload_groups WHERE name = N'Analytics')
  BEGIN;
    SELECT @SQL = 
    'CREATE WORKLOAD GROUP [Analytics]
    WITH(group_max_requests=0
	,importance=Medium
	,request_max_cpu_time_sec=0
	,request_max_memory_grant_percent=25
	,request_memory_grant_timeout_sec=0
	,max_dop='+@MaxDop+')
    USING [Analytics]';
    IF CAST(SERVERPROPERTY('ProductVersion') AS CHAR(2)) >= '13'
      SELECT @SQL = @SQL + ', EXTERNAL [Analytics]';
    EXECUTE (@SQL);
  END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_workload_groups WHERE name = N'AdHoc')
BEGIN;
  SELECT @SQL = 
  'CREATE WORKLOAD GROUP [AdHoc]
    WITH(group_max_requests=0
	,importance=Medium
	,request_max_cpu_time_sec=0
	,request_max_memory_grant_percent=25
	,request_memory_grant_timeout_sec=0
	,max_dop='+@MaxDop+')
    USING [AdHoc]';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_workload_groups WHERE name = N'BatchDay')
BEGIN;
  SELECT @SQL = 
  'CREATE WORKLOAD GROUP [BatchDay]
    WITH(group_max_requests=8
	,importance=Medium
	,request_max_cpu_time_sec=0
	,request_max_memory_grant_percent=25
	,request_memory_grant_timeout_sec=0
	,max_dop='+@MaxDop+')
   USING [Batch]';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_workload_groups WHERE name = N'BatchNight')
BEGIN;
  SELECT @SQL = 
  'CREATE WORKLOAD GROUP [BatchNight]
    WITH(group_max_requests=0
	,importance=Medium
	,request_max_cpu_time_sec=0
	,request_max_memory_grant_percent=25
	,request_memory_grant_timeout_sec=0
	,max_dop='+@MaxDop+')
   USING [Batch]';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_workload_groups WHERE name = N'BatchMaint')
BEGIN;
  SELECT @SQL = 
  'CREATE WORKLOAD GROUP [BatchMaint]
    WITH(group_max_requests=6
	,importance=Medium
	,request_max_cpu_time_sec=0
	,request_max_memory_grant_percent=25
	,request_memory_grant_timeout_sec=0
	,max_dop='+@MaxDop+')
   USING [Batch]';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_workload_groups WHERE name = N'DBA')
BEGIN;
  SELECT @SQL = 
  'CREATE WORKLOAD GROUP [DBA]
    WITH(group_max_requests=0
	,importance=Medium
	,request_max_cpu_time_sec=0
	,request_max_memory_grant_percent=25
	,request_memory_grant_timeout_sec=0
	,max_dop=0)
    USING [DBA]';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_workload_groups WHERE name = N'OLTP')
BEGIN;
  SELECT @SQL = 
  'CREATE WORKLOAD GROUP [OLTP]
    WITH(group_max_requests=0
	,importance=Medium
	,request_max_cpu_time_sec=0
	,request_max_memory_grant_percent=25
	,request_memory_grant_timeout_sec=0
	,max_dop='+@MaxDop+')
    USING [OLTP]';
  EXECUTE (@SQL);
END;

IF NOT EXISTS ( SELECT name FROM sys.resource_governor_workload_groups WHERE name = N'Report')
BEGIN;
  SELECT @SQL = 
  'CREATE WORKLOAD GROUP [Report]
    WITH(group_max_requests=0
	,importance=Medium
	,request_max_cpu_time_sec=0
	,request_max_memory_grant_percent=25
	,request_memory_grant_timeout_sec=0
	,max_dop='+@MaxDop+')
   USING [Report]';
  EXECUTE (@SQL);
END;


IF NOT EXISTS ( SELECT name FROM sys.resource_governor_workload_groups WHERE name = N'System')
BEGIN;
  SELECT @SQL = 
  'CREATE WORKLOAD GROUP [System]
    WITH(group_max_requests=0
	,importance=Medium
	,request_max_cpu_time_sec=0
	,request_max_memory_grant_percent=25
	,request_memory_grant_timeout_sec=0
	,max_dop='+@MaxDop+')
   USING [System]';
  EXECUTE (@SQL);
END;

-- Setup Classifier Tables

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FB_ResourceGovernorClasses]') AND type in (N'U'))
BEGIN;
  CREATE TABLE [dbo].[FB_ResourceGovernorClasses]
  ([Id]                 [int] IDENTITY(1,1) NOT NULL
  ,[AppName]            [nvarchar](256) NULL
  ,[DBName]             [nvarchar](256) NULL
  ,[TimeStart]          [char](5) NOT NULL
  ,[TimeEnd]            [char](5) NOT NULL
  ,[WorkloadGroup]      [nvarchar](256) NOT NULL
    CONSTRAINT [PK_FB_ResourceGovernorClasses] PRIMARY KEY CLUSTERED 
    ([Id] ASC)
    WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON));
	
  CREATE UNIQUE INDEX [IX_FB_ResourceGovernorClasses_Role_AppName_DBName_TimeStart_TimeEnd] ON [dbo].[FB_ResourceGovernorClasses]
  ([AppName] Desc,[DBName] Desc, [TimeStart] Asc, [TimeEnd] Desc)
  INCLUDE([Id], [WorkloadGroup]);
END;

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FB_sysjobnames]') AND type in (N'U'))
BEGIN;
  CREATE TABLE [dbo].[FB_sysjobnames]
  ([job_id] [uniqueidentifier] NOT NULL
  ,[name] [sysname] NOT NULL
  ,[AppName] [nvarchar](256) NOT NULL
  CONSTRAINT [PK_FB_sysjobnames] PRIMARY KEY CLUSTERED 
  ([job_id] ASC));
END;
GO
-- Initial data load of Classifier table

TRUNCATE TABLE [dbo].[FB_ResourceGovernorClasses];

IF NOT EXISTS (SELECT 1 FROM [dbo].[FB_ResourceGovernorClasses] WHERE [AppName] = N'Analytics')
  INSERT INTO [dbo].[FB_ResourceGovernorClasses]
    ([AppName],[DBName],[TimeStart],[TimeEnd],[WorkloadGroup])
    VALUES (N'Analytics',NULL,'00:00','24:00','Analytics');

IF NOT EXISTS (SELECT 1 FROM [dbo].[FB_ResourceGovernorClasses] WHERE [AppName] = N'AdHoc')
  INSERT INTO [dbo].[FB_ResourceGovernorClasses]
    ([AppName],[DBName],[TimeStart],[TimeEnd],[WorkloadGroup])
    VALUES (N'AdHoc',NULL,'00:00','24:00','AdHoc');

IF NOT EXISTS (SELECT 1 FROM [dbo].[FB_ResourceGovernorClasses] WHERE [AppName] = N'Batch' AND [WorkloadGroup] = N'BatchNight')
  INSERT INTO [dbo].[FB_ResourceGovernorClasses]
    ([AppName],[DBName],[TimeStart],[TimeEnd],[WorkloadGroup])
    VALUES (N'Batch',NULL,'18:00','07:00','BatchNight');

IF NOT EXISTS (SELECT 1 FROM [dbo].[FB_ResourceGovernorClasses] WHERE [AppName] = N'Batch' AND [WorkloadGroup] = N'BatchDay')
  INSERT INTO [dbo].[FB_ResourceGovernorClasses]
    ([AppName],[DBName],[TimeStart],[TimeEnd],[WorkloadGroup])
    VALUES (N'Batch',NULL,'07:00','18:00','BatchDay');

IF NOT EXISTS (SELECT 1 FROM [dbo].[FB_ResourceGovernorClasses] WHERE [AppName] = N'BatchMaint' AND [WorkloadGroup] = N'BatchMaint')
  INSERT INTO [dbo].[FB_ResourceGovernorClasses]
    ([AppName],[DBName],[TimeStart],[TimeEnd],[WorkloadGroup])
    VALUES (N'BatchMaint',NULL,'00:00','24:00','BatchMaint');

IF NOT EXISTS (SELECT 1 FROM [dbo].[FB_ResourceGovernorClasses] WHERE [AppName] = N'DBA')
  INSERT INTO [dbo].[FB_ResourceGovernorClasses]
    ([AppName],[DBName],[TimeStart],[TimeEnd],[WorkloadGroup])
    VALUES ('DBA',NULL,'00:00','24:00','DBA');

IF NOT EXISTS (SELECT 1 FROM [dbo].[FB_ResourceGovernorClasses] WHERE [AppName] = N'OLTP')
  INSERT INTO [dbo].[FB_ResourceGovernorClasses]
    ([AppName],[DBName],[TimeStart],[TimeEnd],[WorkloadGroup])
    VALUES (N'OLTP',NULL,'00:00','24:00','OLTP');

IF NOT EXISTS (SELECT 1 FROM [dbo].[FB_ResourceGovernorClasses] WHERE [AppName] = N'Report')
  INSERT INTO [dbo].[FB_ResourceGovernorClasses]
    ([AppName],[DBName],[TimeStart],[TimeEnd],[WorkloadGroup])
    VALUES (N'Report',NULL,'00:00','24:00','Report');

IF NOT EXISTS (SELECT 1 FROM [dbo].[FB_ResourceGovernorClasses] WHERE [AppName] = N'System')
  INSERT INTO [dbo].[FB_ResourceGovernorClasses]
    ([AppName],[DBName],[TimeStart],[TimeEnd],[WorkloadGroup])
    VALUES (N'System',NULL,'00:00','24:00','System');

-- Setup Classifier Function

ALTER RESOURCE GOVERNOR WITH (CLASSIFIER_FUNCTION = NULL);
ALTER RESOURCE GOVERNOR RECONFIGURE;

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FB_ResourceGovernorClassifier]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
  DROP FUNCTION [dbo].[FB_ResourceGovernorClassifier];
GO
  CREATE FUNCTION [dbo].[FB_ResourceGovernorClassifier]()
  RETURNS SYSNAME
  WITH SCHEMABINDING
  AS
  BEGIN;
    DECLARE
     @AppNameBase       nvarchar(256)
    ,@AppName           nvarchar(256)
    ,@DBName            nvarchar(256)
    ,@JobId             uniqueidentifier
    ,@JobName           nvarchar(256)
    ,@RunTime           datetime
    ,@Time              Char(5)
    ,@WorkloadGroup     sysname;

    SELECT 
      @AppNameBase    = APP_NAME()
     ,@DBName         = ORIGINAL_DB_NAME()
     ,@JobId          = CASE WHEN @AppNameBase LIKE 'SQLAgent%Jobstep%' THEN Cast(Convert(binary(16), Substring(@AppNameBase, CHARINDEX('(Job 0x', @AppNameBase) + 5, 34), 1) as uniqueidentifier) END
     ,@AppName        = CASE WHEN @AppNameBase LIKE '.Net SqlClient%' THEN 'OLTP'
                             WHEN @AppNameBase LIKE 'RCSmall%' THEN 'Analytics'
                             WHEN @AppNameBase LIKE 'RTerm.exe' THEN 'Analytics'
                             WHEN @AppNameBase LIKE 'BxlServer.exe' THEN 'Analytics'
                             WHEN @AppNameBase LIKE 'Microsoft Office%' THEN 'AdHoc'
                             WHEN @AppNameBase LIKE 'SQLPS %' THEN 'AdHoc'
                             WHEN @AppNameBase LIKE 'Microsoft SQL Server Management Studio%' THEN 'AdHoc'
                             WHEN @AppNameBase LIKE 'Microsoft%Visual Studio%' THEN 'AdHoc'
                             WHEN @AppNameBase LIKE '%Print%' THEN 'Batch'
                             WHEN @JobId IS NOT NULL AND @DBName = 'msdb' THEN 'System'
                             WHEN @JobId IS NOT NULL THEN 'Batch'
                             WHEN @AppNameBase LIKE '%Report%' THEN 'Report'
                             WHEN @AppNameBase LIKE 'DQ Services%' THEN 'System'
                             WHEN @AppNameBase LIKE 'DatabaseMail%' THEN 'System'
                             WHEN @AppNameBase LIKE 'Microsoft%Windows%' THEN 'System'
                             WHEN @AppNameBase LIKE 'SQLAgent%' THEN 'System'
                             WHEN @AppNameBase LIKE 'SQL Server CEIP%' THEN 'System'
                             WHEN @AppNameBase LIKE 'SQL Server Data Collector%' THEN 'System'
                             ELSE @AppNameBase END
     ,@AppName        = CASE WHEN @AppName <> 'Adhoc' THEN @AppName
                             WHEN IS_SRVROLEMEMBER('sysadmin') = 1 THEN 'DBA'
                             ELSE @AppName END
     ,@RunTime        = Getdate()
     ,@Time           = Convert(Char(5), @Runtime, 14);

    IF @JobId IS NOT NULL
    BEGIN;
      SELECT @AppName = ISNULL(AppName, @AppName)
      FROM dbo.FB_sysjobnames
      WHERE job_id = @JobId;
    END;

    SELECT TOP 1
      @WorkloadGroup = WorkloadGroup
    FROM [dbo].[FB_ResourceGovernorClasses]
    WHERE @AppName   = ISNULL([AppName],@AppName)
      AND @DBName    = ISNULL([DBName],@DBName)
      AND @Time      BETWEEN [TimeStart] AND [TimeEnd]
    ORDER BY [AppName] Desc,[DBName] Desc;

   RETURN @WorkloadGroup;
 END;
GO

-- Setup Backup Jobs proc
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FB_RGBackupJobNames]') AND type = 'P')
  DROP PROCEDURE [dbo].[FB_RGBackupJobNames];
GO
  CREATE PROCEDURE [dbo].[FB_RGBackupJobNames]
  AS
  BEGIN;
    SET NOCOUNT ON;

    WITH BackupJobs AS
     (SELECT job_id,name FROM [msdb].[dbo].[sysjobs] WHERE [name] LIKE 'Backup DB %')
    MERGE [master].[dbo].[FB_sysjobnames] AS TARGET
    USING BackupJobs AS SOURCE
    ON SOURCE.[job_id] = TARGET.[job_id]
    WHEN NOT MATCHED BY TARGET THEN
     INSERT ([job_id], [name], [AppName])
     VALUES (SOURCE.[job_id], SOURCE.[name], 'BatchMaint')
    WHEN NOT MATCHED BY SOURCE AND TARGET.[name] LIKE 'Backup DB %' THEN
     DELETE;

  END;

-- Activate Resource Governor

ALTER RESOURCE GOVERNOR WITH (CLASSIFIER_FUNCTION = [dbo].[FB_ResourceGovernorClassifier]);
IF CAST(SERVERPROPERTY('ProductVersion') AS CHAR(2)) >= '12'
BEGIN;
  DECLARE
    @SQL                        Varchar(max);
  SELECT @SQL = 
  'ALTER RESOURCE GOVERNOR WITH (MAX_OUTSTANDING_IO_PER_VOLUME = DEFAULT);';
  EXECUTE (@SQL);
END;
ALTER RESOURCE GOVERNOR RECONFIGURE;
