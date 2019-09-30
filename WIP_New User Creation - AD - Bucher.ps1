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
$Script:UPN = "$($FirstName).$($LastName)@Macdonald-Johnston.com.au"

# NOTE: Test this section, I am not sure how it will react when assigning the primary address if it already exists.
$Script:Mail = $FirstName + "." + $LastName
$Script:PrimaryDomain = "BucherMunicipal.com.au"
$Script:SecondaryDomains = @("Bucher.com.au", "JDMacdonald.com.au", "MacdonaldJohnston.com.au", "Macdonald-Johnston.com.au", "MJE.com.au")
$Script:ProxyAddresses = @(Get-Addresses -PrimaryDomain $PrimaryDomain -SecondaryDomains $SecondaryDomains)
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

$ConfirmUserCreation = Confirm-NewUserDetails 
    -Firstname $Firstname 
    -Lastname $Lastname 
    -JobTitle $JobTitle 
    -SamAccountName $SAM 
    -EmailAddress $EmailAddress 
    -Password $Password 
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

$Script:NewUser = New-UserAccount 
    -Firstname $Firstname 
    -Lastname $Lastname 
    -SamAccountName $SAM 
    -UPN $UPN 
    -JobTitle $JobTitle 
    -PhoneNumber $PhoneNumber 
    -MirrorUser $MirrorUser 
    -OU $OU 
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
Set-ProxyAddresses -Identity $NewUser -ProxyAddresses $ProxyAddresses

<#
----------
Create the user's home drive
----------
#>
Write-Heading "Creating Home Drive"

$Script:UserFolderDirectory = "\\mjemelfs2\user$"
$Script:Domain = "vicmje"

$HomeDrive = New-HomeDrive 
    -SamAccountName $SAM 
    -Domain $Domain 
    -ParentFolderPath $UserFolderDirectory 
    -FolderName $SAM
    -DriveLetter "H"

<#
----------
Enable the user's mailbox
----------
#>

Write-Heading "Mailbox"

Enable-UserMailbox -Identity $SAM -Alias $EmailAddress

<#
----------
Finishing tasks
----------
#>

Write-Heading "User Creation Complete"