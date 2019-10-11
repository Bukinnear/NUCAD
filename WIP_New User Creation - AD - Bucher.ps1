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
$Script:Password = Get-Password -Password 'rises-zA7N!'

# Pre-windows 2000 Logon name. This is normally what the user will log in with. 
#LDAP Field: SamAccountName
$Script:SAM = $Name.FirstClean + "." + $Name.LastClean

# User logon name/User Principle Name. This will control which domain the user is created under.
# LDAP Field: UserPrincipleName (User Logon Name).
$Script:UPN = "$($Name.FirstClean).$($Name.LastClean)@Macdonald-Johnston.com.au"

$Script:Mail = $Name.FirstClean + "." + $Name.LastClean
$Script:PrimaryDomain = "BucherMunicipal.com.au"
$Script:SecondaryDomains = @()#"Bucher.com.au", "JDMacdonald.com.au", "MacdonaldJohnston.com.au", "Macdonald-Johnston.com.au", "MJE.com.au")

$Addresses = Get-Addresses `
    -MailName $Mail `
    -PrimaryDomain $PrimaryDomain `
    -SecondaryDomains $SecondaryDomains
    
# Ensure that this account does not already exist
If (!(Confirm-AccountDoesNotExist -SamAccountName $SAM))
{
    Write-Space
    Write-Warning "`r`nUser with SAM account of $($SAM) already exists!"
    Write-Heading "Cancelled user creation"
    return
}

$Script:MirrorUser = Get-MirrorUser -UsernameFormat "Firstname Lastname = Firstname.Lastname"
$Script:OU = Get-OU $MirrorUser

$ConfirmUserCreation = Confirm-NewAccountDetails `
    -Firstname $Name.First `
    -Lastname $Name.Last `
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
    -Firstname $Name.FirstClean `
    -Lastname $Name.LastClean `
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
Create the user's home drive
----------
#>

Write-Heading "Creating Home Drive"

New-HomeDrive `
    -SamAccountName $SAM `
    -Domain "vicmje" `
    -HomeDriveDirectory "\\mjemelfs2\user$" `
    -FolderName $SAM -DriveLetter "H"

<#
----------
Enable the user's mailbox
----------
#>

Write-Heading "Mailbox"

Enable-UserMailbox -Identity $SAM -Alias $Mail -Database "01-USERDB" -ExchangeYear "2010"

<#
----------
Finishing tasks
----------
#>

Write-Heading "User Creation Complete"