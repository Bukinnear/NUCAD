Write-Host "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

# Get the script's file path
$Script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Add this path to the modules search directory
if ($env:PSModulePath -notlike "*$($Script:ScriptPath)*")
{
    $env:PSModulePath += (";" + $ScriptPath)
}

Import-Module -Name "NUC-AD" -Force
if (!(Initialize-Module)) { Write-Error "Could not import NUC-AD module" }

#Write-Warning "It is not recommended to create service accounts with this tool."

<#
----------
Get new user details
----------
#>

#Write-Heading "Starting User Creation Process"

$Script:FirstName = "Test"
$Script:LastName = "Automation"
$Script:JobTitle = "Delete Me"
$Script:PhoneNumber = ""
$Script:Password = "rises-zA7N!"
$Script:SAM = "Test.Automation2"
$Script:UPN = "Test.Automation2@peninsuleisure.com.au"
$Script:Mail = "Test.Automation2"
$Script:EmailAddress = "Test.Automation2@peninsuleisure.com.au"
$Script:ProxyAddresses = @()
$Script:MirrorUser = Get-ADUser "test.automation" -Properties *
$Script:OU = Get-OU $MirrorUser
$Script:NewUser = get-aduser "test.automation2" -Properties *

