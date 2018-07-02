Function Set-SessionConfig
{
 Param( [string]$user )
 $account           = New-Object Security.Principal.NTAccount $user
 $sid               = $account.Translate([Security.Principal.SecurityIdentifier]).Value
 
 $config            = Get-PSSessionConfiguration -Name "Microsoft.PowerShell"
 $existingSDDL      = $Config.SecurityDescriptorSDDL
 
 $isContainer       = $false
 $isDS              = $false
 $SecurityDescriptor  = New-Object -TypeName Security.AccessControl.CommonSecurityDescriptor -ArgumentList $isContainer,$isDS, $existingSDDL
 $accessType        = "Allow"
 $accessMask        = 268435456
 $inheritanceFlags  = "none"
 $propagationFlags  = "none"
 $SecurityDescriptor.DiscretionaryAcl.AddAccess($accessType,$sid,$accessMask,$inheritanceFlags,$propagationFlags)
 $SecurityDescriptor.GetSddlForm("All")
} #end Set-SessionConfig
 
# *** Entry Point to script ***  From Ed Wilson, Microsoft Scripting Guy ***
$user               = $($args[0])
$newSDDL            = Set-SessionConfig -user $user
Get-PSSessionConfiguration | ForEach-Object {Set-PSSessionConfiguration -name $_.name -SecurityDescriptorSddl $newSDDL -force }
