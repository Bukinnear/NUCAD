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

# Get the first and last name, and loop until it is approved
do 
{
    Write-Space
    $Script:Firstname = Get-FirstName

    Write-Space
    $Script:Lastname = Get-Lastname

    # Clean the given names
    $ShouldFixCapitalisation = Get-Confirmation "----------`r`nWould you like to fix/standardise the capitalisation of this name?`r`n(Choose yes if you are not sure)"

    $Firstname = Optimize-Name -Name $Firstname -FixCapitalisation $ShouldFixCapitalisation
    $Lastname = Optimize-Name -Name $LastName -FixCapitalisation $ShouldFixCapitalisation

    #Confirm the name is correct
    $ShouldContinue = Confirm-Name -FirstName $Firstname -LastName $Lastname
} while (!$ShouldContinue)

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
WARNING: This may be different from $Mail
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

<#
Mail:
Email 
(Only used below in primary/alias SMTP addresses, this can change later in the script)
LDAP Fields:   
- N/A 

EmailAddress:
Used as the primary SMTP address (this is set with the proxy addresses below)
LDAP Fields:
- Mail (EmailAddress)
- PrimarySmtpAddress (not explicitly set, but by extension of proxy addresses
#>
$Script:Mail = $FirstName+"."+$LastName

# Get Proxy Addresses
for (;;)
{
    $Script:EmailAddress = "$($Mail)@CEDA.com.au"
    $Script:ProxyAddresses = "SMTP:$($EmailAddress)", "smtp:$($Mail)@CEDA.mail.onmicrosoft.com"

    if (Confirm-PrimarySMTPAddress -PrimarySMTP $EmailAddress) 
    {
        break
    }
    else 
    {
        $Script:Mail = Read-Host "Enter the email address (Only enter the text BEFORE the '@' sign)"    
    }
}

$Script:UsernameFormat = "Firstname Lastname = LastnameF"

$Script:MirrorUser = Get-MirrorUser -UsernameFormat $UsernameFormat

$Script:OU = $MirrorUser.DistinguishedName -replace '^cn=.+?(?<!\\),'

# Confirm
if (!(Confirm-NewUserDetails))
{
    Write-Output "`r`n----------`r`nCancelled user creation`r`n----------`r`n"
    return
}