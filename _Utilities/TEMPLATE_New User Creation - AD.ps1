Write-Host "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

<#
----------
Import the New User Creation - AD Module
----------
#>

<#
NOTE: 
This section could replaced with [Import-Module "C:/Path/to/script/NUC-AD" -force ] (no file extension is required in the path),
but the below will allow the script to automatically import the NUC-AD module beside this script no matter where they are placed.
Just don't touch this if you don't need to, and you should be alright.
#>

# Get the script's file path
$Script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Add this path to the modules search directory
if ($env:PSModulePath -notlike "*$($Script:ScriptPath)*")
{
    $env:PSModulePath += (";" + $ScriptPath)
}

Import-Module -Name "NUC-AD" -Force
if (!(Initialize-Module)) 
{ 
    Write-Error "Could not intialize NUC-AD module" 
    return
}

Write-Warning "It is not recommended to create service accounts with this tool."

<#
----------
Get new user details
----------
#>

Write-Heading "Starting User Creation Process"

$Script:Name = Get-Fullname # Note: This variable is a hashtable with firstname, lastname, and cleaned varients thereof
$Script:JobTitle = Get-JobTitle
$Script:PhoneNumber = Get-PhoneNumber
$Script:Password = Get-Password 

# Pre-windows 2000 Logon name. This is normally what the user will log in with. 
#LDAP Field: SamAccountName
$Script:SAM = # $Name.FirstnameClean + "." + $Name.LastnameClean # NOTE: This will vary per client


# User logon name/User Principle Name. This will control which domain the user is created under.
# LDAP Field: UserPrincipleName (User Logon Name).
$Script:UPN = # "$($Name.FirstnameClean).$($Name.LastnameClean)@Domain.com.au" # NOTE: Check what domains the user can be created under in AD

# Ensure that this account does not already exist
If (!(Confirm-AccountDoesNotExist -SamAccountName $SAM))
{
    Write-Warning "`r`nA User with the same name/username already exists!"
    Write-Heading "Cancelled user creation"
    return
}

$Script:Mail = $Name.FirstnameClean + "." + $Name.LastnameClean # Everything that comes before the '@' in a the email address # NOTE: This will vary per client. 
$Script:PrimaryDomain = "Domain.com.au" # The primary domain name # NOTE: This will vary per client
$Script:SecondaryDomains = @("Domain.org", "Domain.com") # Fill this with comma-separated values of any extra domains the user may need in their aliases. NOTE: NOT requried for On-prem Exchange. Leave this blank if that is the case

$Script:Addresses = Get-Addresses -MailName $Mail -PrimaryDomain $PrimaryDomain -SecondaryDomains $SecondaryDomains

$Script:MirrorUser = Get-MirrorUser -UsernameFormat "Firstname Lastname = Firstname.Lastname" # NOTE: This will vary per client
$Script:OU = Get-OU $MirrorUser

$ConfirmUserCreation = Confirm-NewAccountDetails `
    -Firstname $Name.Firstname `
    -Lastname $Name.Lastname `
    -JobTitle $JobTitle `
    -SamAccountName $SAM `
    -EmailAddress $Addresses.PrimarySMTP `
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
    -Firstname $Name.Firstname `
    -Lastname $Name.Lastname `
    -SamAccountName $SAM `
    -UPN $UPN `
    -EmailAddress $Addresses.PrimarySMTP `
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

<#
----------
Populate the account details
----------
#>

Write-Heading "Populating account details."

Set-MirroredProperties -Identity $NewUser.SamAccountName -MirrorUser $MirrorUser
Set-MirroredGroups -Identity $NewUser.SamAccountName -MirrorUser $MirrorUser

# NOTE: O365 ONLY These 3 are NOT TO BE USED if the client has an on-prem exchange. Use Enable-UserMailbox instead
Set-ProxyAddresses -Identity $NewUser.SamAccountName -ProxyAddresses $Addresses.ProxyAddresses
Set-LDAPMail -Identity $NewUser.SamAccountName -PrimarySmtpAddress $Addresses.PrimarySMTP
Set-LDAPMailNickName -Identity $NewUser.SamAccountName -SamAccountName $SAM

<#
----------
Create the user's home drive
----------
#>
# NOTE: Remove this entire section if the client does not have Home drives

Write-Heading "Creating Home Drive"

# NOTE: Remove the comments in this command, or it will fail!
New-HomeDrive `
    -SamAccountName $SAM ` # NOTE: This command will use the user's username for the folder name REMOVE ME
    -Domain "domainname" `
    -HomeDriveDirectory "\\Server\Directory\Path" ` # Root directory path, without the new user's folder name REMOVE ME
    -FolderName $SAM `
    -DriveLetter "H"

<#
----------
Enable the user's mailbox
----------
#>
# NOTE: Remove this entire section if the client does not have an On-prem exchange server
Write-Heading "Mailbox"

Enable-UserMailbox `
    @Name ` # Note: Required to stop the exchange server from using the using the unclean user's name 
    -Identity $NewUser.SamAccountName `
    -Alias $Mail `
    -Database "YourDatabaseName" `
    -ExchangeYear 2016 # the year of the exchange server

<#
----------
Finishing tasks
----------
#>

# Not required for onprem exchange servers
if (Get-Confirmation "Would you like to run a sync to O365?")
{
    Start-O365Sync
}

Write-Heading "User Creation Complete"