Write-Output "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

# Path to new user creation module
$Script:ScriptPath = (split-path -parent $MyInvocation.MyCommand.Definition)

# For debugging purposes
if (($ScriptPath -eq "") -or ($null -eq $ScriptPath))
{
    $ScriptPath = "C:\Code\PS\Testing\New Users"
}

# Import the new user creation module
# NOTE: This needs some work on the failure mode
Import-Module -Name "$($ScriptPath)\NUC-AD" -Force

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

# NOTE: Test this section, I am not sure how it will react when assigning the primary address if it already exists.
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
if (!$ConfirmAccountCreation)
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