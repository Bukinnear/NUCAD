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

Write-Warning "It is not recommended to create service accounts with this tool."

<#
----------
Get new user details
----------
#>

Write-Heading "Starting User Creation Process"

$Fullname = Get-Fullname
$Script:FirstName = $Fullname[0]
$Script:LastName = $Fullname[1]
$Script:JobTitle = Get-JobTitle
$Script:PhoneNumber = Get-PhoneNumber
$Script:Password = Get-Password

# Pre-windows 2000 Logon name. This is normally what the user will log in with. 
#LDAP Field: SamAccountName
$Script:SAM = $FirstName + "." + $LastName

# User logon name/User Principle Name. This will control which domain the user is created under.
# LDAP Field: UserPrincipleName (User Logon Name).
$Script:UPN = "$($FirstName).$($LastName)@peninsulaleisure.com.au"

$Script:Mail = $FirstName + "." + $LastName
$Script:PrimaryDomain = "Peninsulaleisure.com.au"
$Script:SecondaryDomains = @()

$Addresses = Get-Addresses -MailName $Mail -PrimaryDomain $PrimaryDomain -SecondaryDomains $SecondaryDomains
$Script:EmailAddress = $Addresses[0]
$Script:ProxyAddresses = $Addresses[1]

$Script:EmailAddress = $ProxyAddresses[0]

# Ensure that this account does not already exist
If (!(Confirm-AccountDoesNotExist -SamAccountName $SAM))
{
    Write-Warning "`r`nUser with SAM account of $($SAM) already exists!"
    Write-Heading "Cancelled user creation"
    return
}

$Script:MirrorUser = Get-MirrorUser -UsernameFormat "Firstname Lastname = Firstname.Lastname"
$Script:OU = Get-OU $MirrorUser

[bool]$ConfirmUserCreation = Confirm-NewAccountDetails `
    -Firstname $Firstname `
    -Lastname $Lastname `
    -JobTitle $JobTitle `
    -SamAccountName $SAM `
    -EmailAddress $EmailAddress `
    -Password $Password `
    -MirrorUser $MirrorUser

# Confirm user creation
if (!$ConfirmUserCreation)
{
    Write-Heading "Cancelled user creation"
    return
}

<#
----------
Create the account
----------
#>

Write-Heading "Beginning user creation"

$Script:NewUser = New-UserAccount `
    -Firstname $Firstname `
    -Lastname $Lastname `
    -SamAccountName $SAM `
    -UPN $UPN `
    -JobTitle $JobTitle `
    -PhoneNumber $PhoneNumber `
    -MirrorUser $MirrorUser `
    -OU $OU `
    -Password $Password

if ($NewUser)
{    
    Write-Space
    Write-Output "- User Created Successfully."
}
else
{
    Write-Space
    Write-Warning "`r`nThere was an error creating the account. Exiting"
    return
}

Write-Heading "Populating account details."

Set-MirroredProperties -Identity $NewUser -MirrorUser $MirrorUser
Set-MirroredGroups -Identity $NewUser -MirrorUser $MirrorUser

<#
----------
Enable the user's mailbox
----------
#>

Write-Heading "Mailbox"

Enable-UserMailbox -Identity $SAM -Alias $EmailAddress -Database "PARCMELPRIDB"

<#
----------
Finishing tasks
----------
#>

Write-Heading "User Creation Complete"