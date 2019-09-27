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
Import-Module -Name "$($ScriptPath)\NUC-AD" -Force

Write-Warning "It is not recommended to create service accounts with this tool."

Write-Output "`r`n----------`r`nStarting User Creation Process`r`n----------"

$Fullname = Get-Fullname
$Script:FirstName = $Fullname[0]
$Script:LastName = $Fullname[1]

Write-Space
$Script:JobTitle = Get-JobTitle

Write-Space
$Script:PhoneNumber = Get-PhoneNumber

Write-Space
$Script:Password = Get-Password

<#
SAM:
Pre-windows 2000 Logon name
LDAP Fields:
- SamAccountName 

UPN:
User logon name/User Principle Name 
WARNING: This may be different from the email address
LDAP Fields:
- UserPrincipleName (User Logon Name)
#>

$Script:SAM = $LastName+$FirstName[0]

$Script:UPN = "$($FirstName+"."+$LastName)@CEDA.com.au"

# Ensure that this account does not already exist
If (!(Confirm-AccountDoesNotExist -SamAccountName $SAM))
{
    Write-Warning "`r`nUser with SAM account of $($SAM) already exists!"
    Write-Output "`r`n----------`r`nCancelled user creation`r`n----------`r`n"
    return
}

$Script:Mail = $FirstName+"."+$LastName
$Script:PrimaryDomain = "Ceda.com.au"
$Script:SecondaryDomains = @("smtp:$($Mail)@CEDA.mail.onmicrosoft.com")

$Script:ProxyAddresses = @(Get-Addresses -PrimaryDomain $PrimaryDomain -SecondaryDomains $SecondaryDomains)

$Script:UsernameFormat = "Firstname Lastname = LastnameF"
$Script:MirrorUser = Get-MirrorUser -UsernameFormat $UsernameFormat

$Script:OU = Get-OU $MirrorUser

$ConfirmUserCreation = Confirm-NewUserDetails -Firstname $Firstname -Lastname $Lastname -JobTitle $JobTitle -SamAccountName $SAM -EmailAddress $EmailAddress -Password $Password -MirrorUser $MirrorUser

# Confirm user creation
if (!$ConfirmUserCreation)
{
    Write-Output "`r`n----------`r`nCancelled user creation`r`n----------`r`n"
    return
}

Write-Output "`r`n----------`r`Beginning user creation`r`n----------`r`n"

$CreationSuccess = New-UserAccount -Firstname $Firstname -Lastname $Lastname -SamAccountName $SAM -UPN $UPN -JobTitle $JobTitle -PhoneNumber $PhoneNumber -MirrorUser $MirrorUser -OU $OU -Password $Password

if ($CreationSuccess)
{
    Write-Space    
    Write-Host "----------`r`nUser Created Successfully."
}
else
{
    Write-Space
    Write-Warning "`r`nThere was an error creating the account. Exiting"
    return
}

$Script:NewUser = Get-ADUser $SAM

Write-Output "`r`n----------`r`nPopulating account details.`r`n----------`r`n"

Set-MirroredProperties -Identity $NewUser -MirrorUser $MirrorUser

Set-MirroredGroups -Identity $NewUser -MirrorUser $MirrorUser

Set-Addresses -Identity $NewUser -ProxyAddresses $ProxyAddresses -EmailAddress $EmailAddress -SAM $SAM

if (Get-Confirmation "Would you like to run a sync to O365?")
{
    Start-O365Sync
}