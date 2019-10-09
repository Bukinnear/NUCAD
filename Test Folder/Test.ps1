Write-Output "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

# Path to new user creation module
$Script:ScriptPath = (split-path -parent $MyInvocation.MyCommand.Definition)
Import-Module -Name "$($ScriptPath)\NUC-AD" -Force
<#
Import-Module -Name "c:\code\ps\testing\new users\NUC-AD" -Force
#>

Write-Warning "It is not recommended to create service accounts with this tool."

<#
----------
Get new user details
----------
#>

Write-Heading "Starting User Creation Process"

$Script:FirstName = "Test"
$Script:LastName = "Account"
$Script:JobTitle = "Delete Me"
$Script:PhoneNumber = ""
$Script:Password = "rises-zA7N!"
$Script:SAM = ""
$Script:UPN = ""
$Script:Mail = ""
$Script:EmailAddress = "Test.Account@domain.com.au"
$Script:ProxyAddresses = @("SMTP:Test.Account@domain.com.au", "smtp:Test.Account@domain.mail.onmicrosoft.com")
$Script:MirrorUser = Get-ADUser "TestK" -Properties *
$Script:OU = Get-OU $MirrorUser