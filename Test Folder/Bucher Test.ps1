Write-Host "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

# Get the script's file path
$Testpath = Split-path -parent $MyInvocation.MyCommand.Definition
$Script:ScriptPath = $Testpath #Split-Path -Parent $Testpath

# Add this path to the modules search directory
if ($env:PSModulePath -notlike "*$($Script:ScriptPath)*")
{
    $env:PSModulePath += (";" + $ScriptPath)
}

Import-Module -Name "NUC-AD" -Force
if (!(Initialize-Module)) { Write-Error "Could not import NUC-AD module" }


Write-Heading "Starting User Creation Process"

$Script:Name = Get-FullName
$Script:JobTitle = "Delete Me"
$Script:PhoneNumber = "03 9691 0555"
$Script:Password = "rises-zA7N!"
$Script:SAM = $Name.FirstClean + "." + $Name.LastClean
$Script:UPN = "$($Name.FirstClean + "." + $Name.LastClean)@BucherMunicipal.com.au"
$Script:Mail = $Name.FirstClean + "." + $Name.LastClean
$Script:Addresses = Get-Addresses `
    -MailName $Mail `
    -PrimaryDomain "BucherMunicipal.com.au" `
    -SecondaryDomains @()
$Script:MirrorUser = Get-ADUser "jctest" -Properties *
$Script:OU = Get-OU $MirrorUser

# Ensure that this account does not already exist
If (!(Confirm-AccountDoesNotExist -SamAccountName $SAM))
{
    Write-Warning "`r`nUser with SAM account of $($SAM) already exists!"
    Write-Heading "Cancelled user creation"
    return
}

New-ADUser `
    -GivenName $Name.First `
    -Surname $Name.Last `
    -Name "$($Name.First) $($Name.Last)" `
    -DisplayName "$($Name.First) $($Name.Last)" `
    -SamAccountName $SAM `
    -UserPrincipalName $UPN `
    -EmailAddress $Addresses.PrimarySMTP `
    -Description $JobTitle `
    -Title $JobTitle `
    -OfficePhone $PhoneNumber `
    -Department $MirrorUser.Department `
    -Path $OU `
    -Enabled $True `
    -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) `
    -ErrorAction Stop

$Script:NewUser = New-UserAccount `
    -Firstname $FirstName `
    -Lastname $LastName `
    -SamAccountName $SAM `
    -EmailAddress $Addresses.PrimarySMTP `
    -UPN $UPN `
    -JobTitle $JobTitle `
    -PhoneNumber $PhoneNumber `
    -MirrorUser $MirrorUser `
    -OU $OU `
    -Password $Password