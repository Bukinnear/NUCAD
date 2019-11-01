Write-Host "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

# Get the script's file path
$Script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Add this path to the modules search directory
if ($env:PSModulePath -notlike "*$($Script:ScriptPath)*")
{
    $env:PSModulePath += (";" + $ScriptPath)
}

# Import the new user creation module
Import-Module -Name "$($ScriptPath)\NUC-AD" -Force
if (!(Initialize-Module)) { Write-Error "Could not import NUC-AD module" }

Write-Warning "It is not recommended to create service accounts with this tool."

<#
----------
Get new user details
----------
#>

Write-Heading "Starting User Creation Process"

$Script:Name = Get-Fullname
$Script:JobTitle = Get-JobTitle
$Script:PhoneNumber = Get-PhoneNumber
$Script:Password = Get-Password

# Pre-windows 2000 Logon name. This is normally what the user will log in with. 
#LDAP Field: SamAccountName
$Script:SAM = "$($LastName+$FirstName[0])"

# User logon name/User Principle Name. This will control which domain the user is created under.
# LDAP Field: UserPrincipleName (User Logon Name).
$Script:UPN = "$($FirstName+"."+$LastName)@CEDA.com.au"

# Ensure that this account does not already exist
If (!(Confirm-AccountDoesNotExist -SamAccountName $SAM))
{
    Write-Warning "`r`nUser with SAM account of $($SAM) already exists!"
    Write-Heading "Cancelled user creation"
    return
}

$Script:Mail = $FirstName + "." + $LastName
$Script:PrimaryDomain = "Ceda.com.au"
$Script:SecondaryDomains = @("CEDA.mail.onmicrosoft.com")

$Addresses = Get-Addresses -MailName $Mail -PrimaryDomain $PrimaryDomain -SecondaryDomains $SecondaryDomains
$Script:EmailAddress = $Addresses[0]
$Script:ProxyAddresses = $Addresses[1]

$Script:MirrorUser = Get-MirrorUser -UsernameFormat "Firstname Lastname = LastnameF"
$Script:OU = Get-OU $MirrorUser

$ConfirmUserCreation = Confirm-NewAccountDetails `
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
    Write-Warning "There was an error creating the account. Exiting"
    return
}

Write-Heading "Populating account details."

Set-MirroredProperties -Identity $NewUser -MirrorUser $MirrorUser
Set-MirroredGroups -Identity $NewUser -MirrorUser $MirrorUser
Set-ProxyAddresses -Identity $NewUser -ProxyAddresses $ProxyAddresses
Set-LDAPMail -Identity $NewUser -PrimarySmtpAddress $EmailAddress
Set-LDAPMailNickName -Identity $NewUser -SamAccountName $SAM

<#
----------
Finishing tasks
----------
#>

if (Get-Confirmation "Would you like to run a sync to O365?")
{
    Start-O365Sync
}

Write-Heading "User Creation Complete"