Write-Output "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

# Path to new user creation module
$Script:ScriptPath = (split-path -parent $MyInvocation.MyCommand.Definition)
Import-Module -Name "$($ScriptPath)\NUC-AD" -Force

Write-Warning "It is not recommended to create service accounts with this tool."

<#
----------
Get new user details
----------
#>

Write-Heading "Starting User Creation Process"

$Script:FirstName = "Test"
$Script:LastName = "Account"
$Script:JobTitle = "Delete Me"
$Script:PhoneNumber = ""
$Script:Password = "rises-zA7N!"
$Script:SAM = "$($LastName+$FirstName[0])"
$Script:UPN = "$($FirstName+"."+$LastName)@ceda.com.au"
$Script:Mail = $FirstName + "." + $LastName
$Script:EmailAddress = "Test.Account@Ceda.com.au"
$Script:ProxyAddresses = @("SMTP:Test.Account@Ceda.com.au", "smtp:Test.Account@Ceda.mail.onmicrosoft.com")
$Script:MirrorUser = Get-ADUser "TestK" -Properties *
$Script:OU = Get-OU $MirrorUser

# Ensure that this account does not already exist
If (!(Confirm-AccountDoesNotExist -SamAccountName $SAM))
{
    Write-Warning "`r`nUser with SAM account of $($SAM) already exists!"
    Write-Heading "Cancelled user creation"
    return
}

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
return
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

Set-MirroredProperties -Identity $NewUser.DistinguishedName -MirrorUser $MirrorUser
Set-MirroredGroups -Identity $NewUser -MirrorUser $MirrorUser
Set-ProxyAddresses -Identity $NewUser -ProxyAddresses $ProxyAddresses
Set-LDAPMail -Identity $NewUser -PrimarySmtpAddress $EmailAddress
Set-LDAPMailNickName -Identity $NewUser -SamAccountName $SAM
