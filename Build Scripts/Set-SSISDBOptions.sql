-- Copyright FineBuild Team © 2016 - 2017.  Distributed under Ms-Pl License
-- Code based on blog post from Mark Tassin.
-- http://cryptoknight.org/index.php?/archives/1-SSIS-Maintenance-Script.html
-- Some indexes suggested by Mark are now included by Microsoft, this file adds the remaining indexes.

-- Disable Lock Escalation on all tables to assist Cleanup process
DECLARE @SQL NVARCHAR(MAX);
SELECT
 @SQL = COALESCE(@SQL + N'ALTER TABLE [' + SCHEMA_NAME(schema_id) + N'].[' + name + N'] SET (LOCK_ESCALATION=DISABLE);','ALTER TABLE [' + SCHEMA_NAME(schema_id) + '].[' + name + '] SET (LOCK_ESCALATION=DISABLE);')
FROM sys.tables;
EXEC sp_executesql @SQL;

-- Add Indexes to assist Cleanup process

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[internal].[event_message_context]') AND name = N'IX_EventMessageContext_event_message_id_operation_id#FB')
  CREATE NONCLUSTERED INDEX [IX_EventMessageContext_event_message_id_operation_id#FB] ON [internal].[event_message_context]
  (
	[event_message_id] ASC,
	[operation_id] ASC
  )
  INCLUDE (
	[context_id]
  )  WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON);

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[internal].[execution_component_phases]') AND name = N'IX_ExecutionComponentPhases_execution_id#FB')
  CREATE NONCLUSTERED INDEX [IX_ExecutionComponentPhases_execution_id#FB] ON [internal].[execution_component_phases]
  (
	[execution_id] ASC
  )
  INCLUDE (
	[phase_stats_id]
  ) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON);

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[internal].[execution_data_statistics]') AND name = N'IX_ExecutionDataStatistics_execution_id#FB')
  CREATE NONCLUSTERED INDEX [IX_ExecutionDataStatistics_execution_id#FB] ON [internal].[execution_data_statistics]
  (
	[execution_id] ASC
  )
  INCLUDE (
	[data_stats_id]
  ) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON);

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[internal].[execution_data_taps]') AND name = N'IX_ExecutionDataTaps_execution_id#FB')
  CREATE NONCLUSTERED INDEX [IX_ExecutionDataTaps_execution_id#FB] ON [internal].[execution_data_taps]
  (
	[execution_id] ASC
  )
  INCLUDE (
	[data_tap_id]
  ) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON);

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[internal].[execution_parameter_values]') AND name = N'IX_ExecutionParameterValue_execution_id#FB')
  CREATE NONCLUSTERED INDEX [IX_ExecutionParameterValue_execution_id#FB] ON [internal].[execution_parameter_values]
  (
	[execution_id] ASC
  )
  INCLUDE (
	[execution_parameter_id],
   	[object_type],
	[parameter_data_type],
	[parameter_name],
	[parameter_value],
	[sensitive_parameter_value],
	[base_data_type],
	[sensitive],
	[required],
	[value_set],
	[runtime_override]) 
  WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON);

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[internal].[execution_parameter_values]') AND name = N'IX_ExecutionParameterValue_parameter_name#FB')
  CREATE NONCLUSTERED INDEX [IX_ExecutionParameterValue_parameter_name#FB] ON [internal].[execution_parameter_values]
  (
	[parameter_name] ASC
  )
  INCLUDE (
	[execution_parameter_id],
   	[execution_id],
	[object_type],
	[parameter_data_type],
	[parameter_value],
	[sensitive_parameter_value],
	[base_data_type],
	[sensitive],
	[required],
	[value_set],
	[runtime_override]) 
  WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON);

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[internal].[execution_property_override_values]') AND name = N'IX_ExecutionPropertyOverrideValues_execution_id#FB')
  CREATE NONCLUSTERED INDEX [IX_ExecutionPropertyOverrideValues_execution_id#FB] ON [internal].[execution_property_override_values]
  (
	[execution_id] ASC
  )
  INCLUDE (
	[property_id]
  ) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON);

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[internal].[extended_operation_info]') AND name = N'IX_ExtendedOperationInfo_operation_id#FB')
  CREATE NONCLUSTERED INDEX [IX_ExtendedOperationInfo_operation_id#FB] ON [internal].[extended_operation_info]
  (
	[operation_id] ASC
  )
  INCLUDE (
	[info_id]
  ) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON);

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[internal].[operation_os_sys_info]') AND name = N'IX_OperationsOsSysInfo_operation_id#FB')
  CREATE NONCLUSTERED INDEX [IX_OperationsOsSysInfo_operation_id#FB] ON [internal].[operation_os_sys_info]
  (
	[operation_id] ASC
  )
  INCLUDE (
	[info_id],
	[total_physical_memory_kb],
	[available_physical_memory_kb],
	[total_page_file_kb],
	[available_page_file_kb],
	[cpu_count]) 
  WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON);

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[internal].[operations]') AND name = N'IX_operations_end_time#FB')
  CREATE NONCLUSTERED INDEX [IX_operations_end_time#FB] ON [internal].[operations]
  (
	[end_time] ASC
  )
  INCLUDE (
	[operation_id],
	[created_time],
	[status])
  WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON);

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[internal].[object_parameters]') AND name = N'IX_ObjectParameters_project_version_lsn#FB')
  CREATE NONCLUSTERED INDEX [IX_ObjectParameters_project_version_lsn#FB] ON [internal].[object_parameters]
  (
	[project_version_lsn]
  )
  INCLUDE (
	[parameter_id])
  WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON);

GO