# Set-SSISDB
 Param( [string]$HostServer, [string]$dbName, [string]$password )
# Copyright FineBuild Team © 2015.  Distributed under Ms-Pl License
# Based on http://msdn.microsoft.com/en-gb/library/gg471509.aspx
 [Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Management.IntegrationServices")
 $ISNamespace       = "Microsoft.SqlServer.Management.IntegrationServices"

 $sqlConnectString  = "Data Source=$HostServer;Initial Catalog=master;Integrated Security=SSPI;"
 $sqlConnection     = New-Object System.Data.SqlClient.SqlConnection $sqlConnectString
 $SSISService       = New-Object $ISNamespace".IntegrationServices" $sqlConnection

 $catalog           = New-Object $ISNamespace".Catalog" ($SSISService, $dbName, $password)
 $catalog.Create()