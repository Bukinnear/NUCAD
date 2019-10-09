Write-Output "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

# Path to new user creation module
$Script:ScriptPath = split-path -parent $MyInvocation.MyCommand.Definition

# For debugging purposes
if (($ScriptPath -eq "") -or ($null -eq $ScriptPath))
{
    $ScriptPath = "C:\Code\PS\Testing\New Users"
}

# Import the new user creation module
Import-Module -Name "$($ScriptPath)\NUC-AD" -Force
if (!(Initialize-Module)) { return }

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


<#
----------
Enable the user's mailbox
----------
#>

Write-Heading "Mailbox"

Enable-UserMailbox -Identity $SAM -Alias $Mail -Database "PARCMELPRIDB"
