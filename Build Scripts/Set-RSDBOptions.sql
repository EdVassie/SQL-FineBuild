-- Copyright FineBuild Team © 2017 - 2019.  Distributed under Ms-Pl License
-- ReportServer database indexes
USE [$(strRSDBName)]
GO
DECLARE
 @SQLVersion Int
,@SQLText    NVarchar(4000);
SELECT @SQLVersion = cmptlevel FROM master..sysdatabases WHERE name = 'master';

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Event]') AND name = N'IX_Event_ProcessStart_TimeEntered#FB')
	CREATE NONCLUSTERED INDEX [IX_Event_ProcessStart_TimeEntered#FB] ON [dbo].[Event]
	(
		[ProcessStart] ASC,
		[TimeEntered] ASC)
	INCLUDE(EventId)
	WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];

IF EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Event]') AND name = N'IX_Event2')
	DROP INDEX [IX_Event2] ON [dbo].[Event];

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Event]') AND name = N'IX_Event_BatchID_TimeEntered#FB')
	CREATE NONCLUSTERED INDEX [IX_Event_BatchID_TimeEntered#FB] ON [dbo].[Event]
	(
		[BatchID] ASC,
		[TimeEntered] ASC)
	INCLUDE(EventId,EventType,EventData)
	WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Keys]') AND name = N'IX_Keys_Client#FB')
	CREATE NONCLUSTERED INDEX [IX_Keys_Client#FB] ON [dbo].[Keys]
	(
		[Client] ASC)
	INCLUDE(InstallationID)
	WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Notifications]') AND name = N'IX_Notifications_ProcessStart_NotificationEntered#FB')
	CREATE NONCLUSTERED INDEX [IX_Notifications_ProcessStart_NotificationEntered#FB] ON [dbo].[Notifications]
	(
		[ProcessStart] ASC,
		[NotificationEntered] ASC)
	INCLUDE(NotificationID,ProcessAfter,BatchID)
	WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];

IF EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Notifications]') AND name = N'IX_Notifications2')
	DROP INDEX [IX_Notifications2] ON [dbo].[Notifications];

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Notifications]') AND name = N'IX_Notifications_BatchID_NotificationEntered#FB')
BEGIN;
	IF @SQLVersion <= 100 
          SET @SQLText = N'CREATE NONCLUSTERED INDEX [IX_Notifications_BatchID_NotificationEntered#FB] ON [dbo].[Notifications] ' +
			N'( ' +
			N' [BatchID] ASC, ' +
			N'[NotificationEntered] ASC) ' +
			N'INCLUDE(NotificationID, SubscriptionID, ActivationID, ReportID, SnapShotDate, Locale, ProcessStart, Attempt, SubscriptionLastRunTime, DeliveryExtension, SubscriptionOwnerID, IsDataDriven, Version) ' +
			N'WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]';
	IF @SQLVersion > 100 
          SET @SQLText = N'CREATE NONCLUSTERED INDEX [IX_Notifications_BatchID_NotificationEntered#FB] ON [dbo].[Notifications] ' +
			N'( ' +
			N' [BatchID] ASC, ' +
			N'[NotificationEntered] ASC) ' +
			N'INCLUDE(NotificationID, SubscriptionID, ActivationID, ReportID, SnapShotDate, Locale, ProcessStart, Attempt, SubscriptionLastRunTime, DeliveryExtension, SubscriptionOwnerID, IsDataDriven, Version, ReportZone) ' +
			N'WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]';
	EXEC (@SQLText);
END;

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Users]') AND name = N'IX_Users_Sid_AuthType#FB')
	CREATE NONCLUSTERED INDEX [IX_Users_Sid_AuthType#FB] ON [dbo].[Users]
	(
		[Sid] ASC,
		[AuthType] ASC
	)
	INCLUDE ( 	[UserID]) 
        WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[Users]') AND name = N'IX_Users_UserName_AuthType#FB')
	CREATE NONCLUSTERED INDEX [IX_Users_UserName_AuthType#FB] ON [dbo].[Users]
	(
		[UserName] ASC,
		[AuthType] ASC
	)
	INCLUDE ( 	[UserID]) 
        WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];

--
-- ReportServer TempDB database indexes
USE [$(strRSDBName)TempDB]
GO

