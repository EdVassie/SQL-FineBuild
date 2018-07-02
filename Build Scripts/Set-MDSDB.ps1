# Set-MDSDB
 Param( [string]$dllPath, [string]$instance, [string]$dbName, [string]$account )
# Copyright FineBuild Team © 2017.  Distributed under Ms-Pl License
Import-Module -Name $dllPath
$server = Get-MasterDataServicesDatabaseServerInformation -ConnectionString 'Data Source=$instance;Initial catalog=;Integrated Security=True;User ID=;Password='; 
New-MasterDataServicesDatabase -Server $server -DatabaseName '$dbName' -AdminAccount '$account';