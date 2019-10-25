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
    -SecondaryDomains @("Bucher.com.au", "JDMacdonald.com.au", "MacdonaldJohnston.com.au", "Macdonald-Johnston.com.au", "MJE.com.au")
$Script:MirrorUser = Get-ADUser "jctest" -Properties *
$Script:OU = Get-OU $MirrorUser

# Ensure that this account does not already exist
If (!(Confirm-AccountDoesNotExist -SamAccountName $SAM))
{
    Write-Warning "`r`nUser with SAM account of $($SAM) already exists!"
    Write-Heading "Cancelled user creation"
    return
}

$Script:NewUser = New-UserAccount `
    -Firstname $Name.First `
    -Lastname $Name.Last `
    -SamAccountName $SAM `
    -UPN $UPN `
    -EmailAddress $Addresses.PrimarySMTP `
    -JobTitle $JobTitle `
    -PhoneNumber $PhoneNumber `
    -MirrorUser $MirrorUser `
    -OU $OU `
    -Password $Password

Set-MirroredProperties -Identity $NewUser -MirrorUser $MirrorUser
Set-MirroredGroups -Identity $NewUser -MirrorUser $MirrorUser
set-ldapmail -Identity $NewUser -PrimarySMTPAddress $Addresses.PrimarySMTP

return

Enable-UserMailbox -Identity "Test.McAccount" -Alias "Test.McAccount" -Database "01-USERDB" -ExchangeYear "2010"