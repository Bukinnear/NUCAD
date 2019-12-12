Write-Output "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

# Path to new user creation module
$Script:ScriptPath = (split-path -parent $MyInvocation.MyCommand.Definition)
Import-Module -Name "$($ScriptPath)\NUC-AD" -Force

$Identity = Get-ADUser "test.user1"
$Groups = @("All Kiandra", "Melbourne Staff")

Add-ADPrincipalGroupMembership -Identity $Identity -MemberOf "All Kiandra", "Melbourne Staff"

<#
"CN=Melbourne Staff,OU=Distribution Groups,OU=Kiandra,DC=kiandra,DC=local"
"CN=All Kiandra,OU=Security Groups,OU=Kiandra,DC=kiandra,DC=local"
#>