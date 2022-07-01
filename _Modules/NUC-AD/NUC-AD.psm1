$LogPath = 'C:\temp\UserCreationADLog.txt'
if (!(Get-ChildItem $LogPath -ErrorAction SilentlyContinue)) 
{
    $NULL = New-Item $LogPath -ItemType File
}
# Import-Module -Name "C:\Code\Powershell\_Modules\NUC-AD" -Force

Add-Type -TypeDefinition @"
   public enum LogTypes
   {
      DEBUG,
      WARNING,
      ERROR
   }
"@

<#
.SYNOPSIS
    Gets a true/false response from the user
.DESCRIPTION
    Writes the provided text to console, and expects the user to provide a yes/no response. Returns 1 or 0 for true and false, respectively
.EXAMPLE
    PS C:\> Get-Confirmation "Do you want to continue?"
    Send the text "Do you want to continue?" to console and prompts the user to provide a yes/no response. Returns 1 or 0 (true or false)
.PARAMETER Message
    The text to write out/prompt the user with
.INPUTS
    String
.OUTPUTS
    Boolean
.NOTES    
#>
function Get-Confirmation
{
    param(
        # Message to prompt
        [Parameter(
            Mandatory=$true,
            Position=0)]
        [string]
        $Message
    )
    :ConfirmationLoop for (;;)
    {
        [String] $msgConfimation = ""
        try
        {
            $msgConfimation = [System.Windows.MessageBox]::Show($Message,'Confirm','YesNo','Question')
        }
        catch 
        {
            $msgConfimation = Read-Host "`r`n$($Message)`r`n`r`nY/N"
        }

        $msgConfimation = $msgConfimation.ToLower()
        $msgConfimation = $msgConfimation.Trim()
    
        switch ($msgConfimation)
        {
            {'yes','y' -icontains $_}{return $true}
            {'no','n' -icontains $_}{return $null}
            default
            {
                Write-Warning "I could not understand your response. Please reply with either yes or no."
                continue
            }            
        }
    }
}

<#
.SYNOPSIS
    Writes out the details of the provided error. Can be directed to console, or console and file.
.DESCRIPTION
    Writes out the provided error provided in CaughtError. LogType can be specified as Error, Warning, or Debug. 
    LogString will provide optional, contextual infomation.
    The log file will be directed to C:\temp\UserCreationLogs\UserCreationADLog.txt
    Does not write to file by default.
.EXAMPLE
    PS C:\> Try {
        Throw "Error"
    }
    catch
    {
        Write-NewestErrorMessage -LogType Error -CaughtError $_ -LogToFile $True -LogString "This command must be run as administrator."
    }
    
    Best used in a try/catch loop (for easy access to the most recent error). This will display error text, the category, and the text provided in LogString.
    The same information will additionally be written to file.
.PARAMETER LogType
    The severity of the information being written. Valid options are Error, Warning, and Debug
.PARAMETER CaughtError
    The error variable you want to write the details of. Can optionally be omitted to provide only infomation provided by LogString.
.PARAMETER LogToFile
    Whether or not this error should be written to file. Accepts a boolean value
.PARAMETER LogString
    The optional, contextual message to provide alongside the error details. Not required, but strongly recommended.
.INPUTS
    LogTypes    
    Boolean
    String
#>
function Write-NewestErrorMessage
{
    param(
        [Parameter(Mandatory=$true)]
        [LogTypes] $LogType, 

        # The caught error
        [Parameter()]
        [System.Management.Automation.ErrorRecord]
        $CaughtError,

        # Wheter or not to output to file. Default is true
        [Parameter()]
        [bool]
        $LogToFile = $false,
            
        [Parameter()]
        [String] 
        $LogString    
    )        

    Switch ($LogType.ToString())
    {
        "DEBUG"
        {
            $LogColour = [System.ConsoleColor]::White            
        }
        "WARNING"
        {
            $LogColour = [System.ConsoleColor]::Yellow
        }
        "ERROR"
        {
            $LogColour = [System.ConsoleColor]::Red
        }
    }

    Write-Host -ForegroundColor $LogColour "`r`n$($LogType): $($LogString)"
    Write-Host "`r`nFull Details:"
    Write-Host -ForegroundColor $LogColour "$($LogType.ToString())`r`nCategory: $($CaughtError.CategoryInfo.Category)`r`nMessage: $($CaughtError.Exception.Message)`r`n"

    if ($LogToFile -and (Get-ChildItem $LogPath -ErrorAction SilentlyContinue))
    {
        $LogFileText = "$(Get-Date -Format "yyyy/MM/dd | HH:mm:ss") | $($LogType.ToString()) | $($LogString) | Category: $($CaughtError.CategoryInfo.Category) | Message: $($CaughtError.Exception.Message)"
        Out-File -FilePath $LogPath -Append -InputObject $LogFileText
    }
}

# 
<#
.SYNOPSIS
    Needs to be run before anything else is called. Returns true if successful, and false if failed.
.DESCRIPTION
    Needs to be run before anything else is called. Returns true if successful, and false if failed.
.EXAMPLE
    PS C:\> if (!Initialize-Module) {
        # Stop the script
    }

    If the initialization fails, provides the opportunity to abort any future actions
.OUTPUTS
    Boolean
#>
function Initialize-Module
{
    param (
        # Specify the year version of Exchange if you intend to use it
        [Parameter(
            Mandatory=$true,
            ParameterSetName='Exchange'
        )]
        [int]
        $Exchange,

        # Blank variable for when exchange is not specified
        [Parameter(
            Mandatory=$false,
            ParameterSetName='NoSet'
        )]
        [int]
        $NoSet
    )
    
    # Check if we are running as admin
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    If (!$currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
    {
        Write-Warning "No admin priviledges detected.`r`n"
        if (!(Get-Confirmation "WARNING: This script is not running as admin, you will probably be unable to create the user account.`r`n`r`nAre you sure you want to continue?"))
        {
            return $false
        }    
    }

    if ($PSCmdlet.ParameterSetName -eq "Exchange")
    {
        if (!(Import-ExchangeSnapin -ExchangeYear $Exchange))
        {
            Write-Warning "Could not import the Exchange snap-in`r`n"

            if (!(Get-Confirmation "WARNING: Could not import the exchange snap-in. You may not be able to create the user mailbox.`r`n`r`nAre you sure you want to continue?"))
            {
                return $false
            } 
        }
    }

    # Import AD Module
    try
    {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not import Active Directory Module. Aborting."
        return $false
    }
    return $true
}

<#
.SYNOPSIS
    Imports the Exchange Server Management Snapin corresponding to the year provided in $ExchangeYear. Accepts 2010, or 2013
.DESCRIPTION
    Imports the Exchange Server Management Snapin corresponding to the year provided in $ExchangeYear. Accepts 2010, or 2013
.EXAMPLE
    PS C:\> Import-ExchangeSnapin "2010"
    Imports the 2010 Exchange Management Snapin
.INPUTS
    String
.PARAMETER ExchangeYear
    The version of Exchange module that will be imported. Accepts "2010", or "2013"
#>
function Import-ExchangeSnapin 
{
    param (
        # The version of Exchange that we will be importing
        [Parameter(
            Mandatory=$true
        )]        
        [Int]
        $ExchangeYear
    )

    # Validate the input year
    if (2007, 2010, 2013, 2016, 2019 -notcontains $ExchangeYear)
    { return $false }
    
    switch ($ExchangeYear)
    {
        2007
        {
            try
            {
                Write-Space
                Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
                return $true
            }
            catch
            {
                Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not import Exchange 2007 Management module."
                return $false                    
            }
        }
        2010
        {  
            try
            {
                add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction Stop
                return $true
            }
            catch
            {
                if ($_.Exception -like "*Microsoft.Exchange.Management.PowerShell.E2010 because it is already added*") { return $true }
                else
                {
                    Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not import Exchange 2010 Management module."    
                    return $false        
                }
            }
        }
        { 2013, 2016, 2019 -contains $_ }
        {
            try
            {
                Write-Space
                Add-PSSnapin Microsoft.Exchange.Management.Powershell.SnapIn
                return $true
            }
            catch
            {
                Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not import Exchange Management module."
                return $false                    
            }
        }
        Default 
        {
            Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not import exchange module - no valid year specified"
            return $false
        }
    }
    # if no conclusion was reached until now
    return $false
}

<#
.SYNOPSIS
    Writes a new line to console
.DESCRIPTION
    Writes a new line to console. Used to easily space console output.
.EXAMPLE
    PS C:\> Write-Space

    PS C:\>    
#>
function Write-Space
{
    Write-Host ""
}

<#
.SYNOPSIS
    Writes out the provided text as a heading banner
.DESCRIPTION
    Writes out the provided text as heading banner with extra lines, and formatted line breaks before and after
.EXAMPLE
    C:\> Write-Heading "Example"

    ----------
    Example
    ----------

    C:\>
.PARAMETER Heading    
    Text to display as a heading
.INPUTS
    String
#>
function Write-Heading
{
    param (
        # The heading text
        [Parameter(
            Mandatory=$true,
            Position=0
        )]
        [String]
        $Heading
    )

    Write-Host "`r`n----------`r`n$($Heading)`r`n----------`r`n"
}

<#
.SYNOPSIS
    Prompt the user to provide a first name 
.DESCRIPTION
    Prompts the user to provide a first name. Outputs the string value provided by the user
.EXAMPLE
    PS C:\> $Name = Get-FirstName
    ----------
    Please enter a First Name: John
    
    PS C:\> $Name
    John
    PS C:\>
.OUTPUTS
    String
#>
function Get-FirstName
{
    [string]$FirstName = Read-Host "----------`r`nPlease enter a First Name"
    return $FirstName
}

<#
.SYNOPSIS
    Prompt the user to provide a last name 
.DESCRIPTION
    Prompts the user to provide a last name. Outputs the string value provided by the user
.EXAMPLE
    PS C:\> $Name = Get-LastName
    ----------
    Please enter a Last Name: John
    
    PS C:\> $Name
    John
    PS C:\>
.OUTPUTS
    String
#>
function Get-Lastname
{
    [string]$LastName = Read-Host "----------`r`nPlease enter a Last Name"
    return $LastName
}

# Get the first and last name, and loop until it is approved. Returns an array containing the First and Last name
<#
.SYNOPSIS
    Prompts the user to provide a first and last name, and returns the result in a hashtable. Names can optionally be cleaned
.DESCRIPTION
    Prompts the user to provide a first and last name, and returns the result in a hashtable. Cleaned names will be provided as well
    The user can also specify whether they want the capitalization fixed as well.

    Cleaned names will have folowing special characters removed:
    ' " ’ ‘ ‛ / \ ; : ( ) [ ] ! @ $ % ^ & * ` ~ . and ‚
    
    The hastable contains the following values:
    Firstname
    Lastname
    FirstnameCleaned
    LastnameCleaned
.EXAMPLE
    PS C:\> Get-FullName

    ----------
    Please enter a First Name: joHN

    ----------
    Please enter a Last Name: d'oe

    ----------
    Would you like to fix/standardise the capitalisation of this name?
    (Choose yes if you are not sure)

    Y/N: y

    ----------
    Is this name correct?

    John D'oe

    Y/N: y

    Name                           Value
    ----                           -----
    FirstnameClean                 John
    Firstname                      John
    LastnameClean                  Doe
    Lastname                       D'oe
.INPUTS
    Boolean
.OUTPUTS
    Hashtable
#>
function Get-FullName
{
    param (
        # Choose whether to prompt for, and fix capitalisation
        [Parameter()]
        [bool]
        $PromptForCapitalisation=$true
    )

    for (;;)
    {
        Write-Space
        $Firstname = Get-FirstName

        Write-Space
        $Lastname = Get-Lastname

        if ($PromptForCapitalisation)
        {
            if (Get-Confirmation "----------`r`nWould you like to fix/standardise the capitalisation of this name?`r`n(Choose yes if you are not sure)")
            {
                $Firstname = Get-CapitalisedName $Firstname
                $Lastname = Get-CapitalisedName $Lastname                
            }
        }

        # Clean the given names
        $FirstnameClean = Get-CleanedName -Name $Firstname
        $LastnameClean = Get-CleanedName -Name $LastName

        # Confirm the name is correct
        If (Confirm-Name -FirstName $Firstname -LastName $Lastname)
        {
            return @{
                Firstname = $Firstname
                Lastname = $Lastname
                FirstnameClean = $FirstnameClean
                LastnameClean = $LastnameClean
            }
        }
    }
}

# Cleans the given string by converting to lowercase and removing leading/trailing white space, and illegal characters
function Get-CleanedName
{
    param (
        # Name to clean
        [Parameter(
            Mandatory=$true
            )]
        [string]
        $Name
    )
    
    # illegal characters
    $CharList = "`'", "`"", "’", "‘", "‛", "/", "\",";", ":", "(", ")", "[", "]", "!", "@", "$", "%", "^", "&", "*", "``", "~", ".", "‚"
    
    # Remove illegal characters
    foreach ($Char in $CharList)
    {
        $Name = $Name.Replace($Char, "")
    }

    #Remove white spaces
    $Name = $Name.Trim()
    $Name = $Name.ToLower()
    
    return $Name
}

# Capitalises the first letter, if user approves
function Get-CapitalisedName
{
    param (
        # Name to capitalise
        [Parameter(
            Mandatory=$true
            )]
        [string]
        $Name
    )

    $Name = (Get-Culture).TextInfo.ToTitleCase($Name.ToLower())

    return $Name    
}

# Prompt the user to provide a job title/description for the new account
Function Get-JobTitle
{
    Write-Space
    return Read-Host "----------`r`nPlease enter a Job Description"
}

# Prompt the user to provide a phone number for the new account
function Get-PhoneNumber
{
    Write-Space
    $PhoneNumber = Read-Host "----------`r`nPlease enter a phone number (Leave blank for none)"        
    
    if ($PhoneNumber)
    {
        $PhoneNumber = $PhoneNumber.Trim()
    }

    return $PhoneNumber
}

function Get-MobileNumber
{
    Write-Space
    $MobileNumber = Read-Host "----------`r`nPlease enter a mobile number (Leave blank for none)"
    if ($MobileNumber)
    {
        $MobileNumber = $MobileNumber.Trim()
    }

    return $MobileNumber
}

# Prompt the user to provide a password for the new account
function Get-Password
{
    Write-Host -ForegroundColor Yellow "`r`nNOTE: Please don't use a 'default' password (consider auto-generating it?)`r`n"
    
    return Read-Host "Please enter a Password" -AsSecureString
}

<#
.SYNOPSIS
    Generates and returns the proxy addresses of the user from the mail name and domains provided. 
.DESCRIPTION
    Generates and returns the proxy addresses of the user from the mail name, and domains provided. 
    Returns a hash map with values PrimarySMTP and ProxyAddresses. 

    PrimarySMTP is an email address generated by concatonating the Mailname and PrimaryDomain.
    Proxyaddresses is an array of addresses generated by concatonating the Mailname and each of the domains listed in Secondary Domains
.EXAMPLE
    PS C:\> $Addresses = Get-Addresses -MailName "Test.Account" -PrimaryDomain "Kiandra.com.au" -SecondaryDomains @("Kiandra.com", "Kiandra.mail.onmicrosoft.com")
    
    PS C:\> $Addresses.PrimarySMTP
    Test.Account@Kiandra.com.au

    PS C:\> $Addresses.ProxyAddresses
    Test.Account@Kiandra.com
    Test.Account@Kiandra.mail.onmicrosoft.com

.PARAMETER MailName
    The name to be used in front of the domain/s
.PARAMETER PrimaryDomain
    # Domain to be used in the primary SMTP Address/Email
.PARAMETER SecondaryDomains
    # Array of domains to be added as aliases/proxy addresses. This MUST be an array, not just a string eg. @("Example.com.au", "Example.com")
.INPUTS
    String
    String
    String
.OUTPUTS
    Hashtable
#>
# Generates and returns the proxy addresses of the user from the mail name, and domains provided. 
function Get-Addresses
{
    param (
        # The name to be used in front of the domain/s
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $MailName,

        # Domain to be used in the primary SMTP Address/Email field
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $PrimaryDomain,

        # Array of domains to be added as aliases/proxy addresses. This MUST be an array, not just a string eg. @("Example.com.au", "Example.com")
        [Parameter()]
        [string[]]
        $SecondaryDomains = @()
    )
    
    for (;;)
    {        
        $PrimaryAddress = "$($MailName)@$($PrimaryDomain)"
        $ProxyAddresses = @("SMTP:$($PrimaryAddress)")
        
        foreach ($Domain in $SecondaryDomains)
        {
            $ProxyAddresses += "smtp:$($MailName)@$($Domain)" 
        }

        $ReturnValue = @{
            PrimarySMTP = $PrimaryAddress
            ProxyAddresses = $ProxyAddresses
        }

        if (Confirm-PrimarySMTPAddress -PrimarySMTP $PrimaryAddress) 
        {
            return $ReturnValue
        }
        else 
        {
            $MailName = Read-Host "Enter the email address (Only enter the text BEFORE the '@' sign)"    
        }
    }
}

# Gets the user specified by the username.
function Get-MirrorUser 
{
    param (
        # The format of the username to be used as a text prompt
        [Parameter()]
        [string]
        $UsernameFormat
    )

    for (;;)
    {
        if ($null -ne $UsernameFormat)
        {
            $Prompt = "`r`n(Username format: $($UsernameFormat))"
        }
        
        $MirrorName = Read-Host "`r`n----------`r`nPlease enter a user to mirror$($Prompt)"
        $MirrorName = $MirrorName.Trim()
        
        try 
        {
            $User = Get-ADUser -Identity $MirrorName -Properties *

            Write-Host "`r`nLocated mirror account:`r`n`r`nAccount`r`n$($User.SAMAccountName)`r`n`r`nName`r`n$($User.Name)`r`n`r`nEmail`r`n$($User.EmailAddress)"

            if (Get-Confirmation "Is this the correct account?")
            {
                return $User
            }            
        }
        catch 
        {
            Write-Space
            if (Get-Confirmation "Could not locate user account - would you like to search for an account?")
            {
                Search-UserAccounts
            }
            else 
            {
                Write-NewestErrorMessage -LogType WARNING -CaughtError $_ -LogString "Could not find any user with that username."                
            }
            continue
        }
    }

    # Should never get here, but just in case
    throw "No mirror user was found"
}

function Get-Manager
{
    for (;;)
    {
        $ManagerName = (Read-Host "`r`n----------`r`nPlease the manager's username").Trim()

        if (!$ManagerName)
        {
            if (Get-Confirmation "Would you like to set the user's manager as 'None'?")
            {
                return $null
            }
            else 
            {
                continue
            }
        }
        
        try 
        {
            $manager = Get-ADUser -Identity $ManagerName -Properties *
            if ($manager -and !$manager.count)
            {
                return $manager.DistinguishedName
            }
            else 
            {
                throw
            }
        }
        catch 
        {
            Write-Space
            if (Get-Confirmation "Could not locate user account - would you like to search for an account?")
            {
                Search-UserAccounts
            }
        }
    }
}

# Gets the OU of the provided user
function Get-OU 
{
    param (
        # User identity to get the OU from
        [Parameter(
            Mandatory=$true
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]
        $MirrorUser
    )

    return $MirrorUser.DistinguishedName -replace '^cn=.+?(?<!\\),'
}

# Checks that the provided account name exists. This will wait and check again for a set number of times if it cannot be immediately found. Returns the user if found.
function Get-NewAccount
{
    param (
        # Username of the required account
        [Parameter(
            Mandatory=$true,
            Position=0
        )]
        [string]
        $SamAccountName,

        # Number of times to try, with each try taking a minimum of 3 seconds. Default is 20 tries
        [Parameter()]
        [int]
        $AttemptCount = 20
    )

    Write-Space
    Write-Host "Waiting for account to become available..."

    for ($i = 0; $i -lt $AttemptCount; $i++)
    {
        try 
        {
            $User = Get-ADUser $SamAccountName -Properties *
            return $User
        } 
        catch 
        { 
            start-sleep -Seconds 3
            continue
        }
    }

    return $null
}

function Search-UserAccounts 
{
    Write-Host "`r`n----------"

    for (;;)
    {        
        $Search = Read-Host "Enter a name, or part of a name to search for. leave it blank to continue`r`n"
    
        if ("" -eq $Search) { return }
        
        try 
        {
            $Results = @(Get-ADUser -Filter "name -like '*$($Search)*'" -Properties SamAccountName, Name, DisplayName, EmailAddress | select Name, DisplayName, SamAccountName, EmailAddress)
        } 
        catch 
        { 
            Write-Space 
        }
        
        if ($Results)
        {
            Write-Host "`r`n----------"
            $Results | Format-Table | Out-String | % {Write-Host $_}
            Write-Host "----------`r`n"
        }
    }
}

# Prompts the user to confirm the given first and last name. Returns true if they approve
function Confirm-Name
{
    param (
        # First name to confirm
        [Parameter(
            Mandatory=$true)]
        [string]
        $FirstName,

        # First name to confirm
        [Parameter(
            Mandatory=$true)]
        [string]
        $LastName
    )
    
    return (Get-Confirmation "----------`r`nIs this name correct?`r`n`r`n$($FirstName) $($LastName)")
}

function Confirm-Username
{
    param(
        # Sam Account Name
        [Parameter(Mandatory=$true)]
        [ref]
        $SamAccountName,

        # User Principal Name
        [Parameter(Mandatory=$true)]
        [ref]
        $UPN
    )

    for (;;)
    {
        $ConfirmMessage = "----------`r`nThe username has been set to:`r`n`r`nLogon Name (UPN): $($UPN.Value)`r`nPre-Windows 2000 (SAM): $($SamAccountName.Value)`r`n`r`nIs this correct?"

        if (Get-Confirmation $ConfirmMessage)
        {
            if (Confirm-AccountDoesNotExist -SamAccountName $SamAccountName.Value)
            {
                return $true
            }
        }

        if ($UPN.Value.IndexOf('@') -eq -1)
        {
            throw "No domain could be found in User's UPN."
        }

        $ReturnValue = (Read-Host "`r`nPlease enter the new User Logon Name (UPN) (without the '@domain').`r`nleave blank to leave unchanged`r`n").trim()

        if ($ReturnValue)
        { 
            $UPN.Value = $ReturnValue + $UPN.Value.Substring($UPN.Value.IndexOf('@'))
        }
        
        if (Get-Confirmation "Would you like to set the Pre-Windows 2000 name to the same? (select 'no' to enter a new one")
        {
            $SamAccountName.Value = $ReturnValue
        }
        else
        {
            $ReturnValue = (Read-Host "`r`nPlease enter the Pre-Windows 2000 username (SAM).`r`nLeave blank to leave unchanged`r`n").trim()
        }

        if ($ReturnValue)
        {
            $SamAccountName.Value = $ReturnValue
        }
    }
}

# Returns true if the account does not already exist
function Confirm-AccountDoesNotExist 
{
    param (
        # Parameter help description
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $SamAccountName
    )

    try 
    {
        $Test = get-ADUser -identity $SamAccountName
    } catch {}

    if ($Test)
    {
        Write-Warning "`r`nA User with the same name/username already exists!"
        return $false
    }
    else 
    {
        return $true
    }
}

# Prompts the user to confirm the given SMTP address. Returns true if they approve.
function Confirm-PrimarySMTPAddress
{
    param(
        # The primary SMTP Address
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $PrimarySMTP
    )

    return Get-Confirmation "----------`r`nThe primary SMTP/email address as been set to:`r`n`r`n$($PrimarySMTP)`r`n`r`nIs this correct?"
}

function Confirm-Manager 
{
    param (
        # User's account identity
        [Parameter(Mandatory=$true)]
        [string]
        $Identity
    )

    if ($null -eq $Identity -or $Identity.Trim() -eq "")
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogString "The identity `"$($Identity)`" is invalid in the `'Confirm-Manager`' command."
        return
    }

    $manager = $null
    try 
    {
        $user = Get-ADUser -identity $Identity -Properties Manager
        if ($user.Manager)
        {
            $manager = Get-ADUser $user.Manager
        }
    } 
    catch 
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogString "Could not get the new user's account to set the manager."
    }

    for (;;)
    {
        if (!($manager))
        {
            $ConfirmationMessage = "No manager has been set. Would you like to set one?"
        }
        else 
        {
            $ConfirmationMessage = "User's manager has been set to `"$($manager.Name)`". Is this correct?"
        }
        
        if (!(Get-Confirmation $ConfirmationMessage))
        {
            return
        }
        else 
        {
            $new_manager = Get-Manager
            if (!$new_manager)
            {
                Set-ADUser -Identity $user -clear manager
                Write-Space
                Write-Host "User's manager has been cleared."
                return
            }
            else 
            {
                $new_manager = get-ADUser -Identity $new_manager -Properties *
            }
        }
        
        Write-Host "`r`nLocated new manager account:`r`n`r`nAccount`r`n$($new_manager.SAMAccountName)`r`n`r`nName`r`n$($new_manager.Name)`r`n`r`nEmail`r`n$($new_manager.EmailAddress)"

        if (Get-Confirmation "Is this the correct account?")
        {
            try
            {
                Set-ADUser -Identity $user -Manager $new_manager.DistinguishedName
                Write-Space
                Write-Host "User's manager has been set!"
                return
            }
            catch
            {
                Write-NewestErrorMessage -LogType WARNING CaughtError $_ -LogString "Something went wrong while setting the user's manager!"
                
            }
        } 
    }
}

# Prompts the user to confirm the new account details
function Confirm-NewAccountDetails 
{
    param (
        [Parameter(
            Mandatory=$true)]
        [string]
        $Firstname,

        [Parameter(
            Mandatory=$true)]
        [string]
        $Lastname,

        [Parameter(
            Mandatory=$true)]
        [string]
        $SamAccountName,

        [Parameter()]
        [string]
        $JobTitle,

        [Parameter(
            Mandatory=$true)]
        [string]
        $EmailAddress,

        [Parameter()]
        [SecureString]
        $Password,

        [Parameter(
            Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADUser]
        $MirrorUser

    )

    $ConfirmationMessage = "----------`r`nYou are about to create an account with the following details:`r`n`r`nFirst Name: $($Firstname)`r`n`r`nLast Name: $($Lastname)`r`n`r`nJob Title: $($JobTitle)`r`n`r`nUsername: $($SamAccountName)`r`n`r`nEmail Address: $($EmailAddress)`r`n`r`nAccount to Mirror: $($MirrorUser.DisplayName)`r`n`r`nDo you wish to proceed?"
    return (Get-Confirmation $ConfirmationMessage)
}

# Creates a new user from the parameters provided. Returns true if the account was created successfully, and false if not.
function New-UserAccount 
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # First name
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Firstname,

        # Last name
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Lastname,

        # Username
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $SamAccountName,

        # Username, and domain
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $UPN,

        # Email address
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $EmailAddress,

        # Job title/account description (Optional)
        [Parameter()]
        [string]
        $JobTitle,

        # The account description - uses Job title by default
        [Parameter()]
        [string]
        $Description = $JobTitle,

        # Phone number (Optional)
        [Parameter()]
        [string]
        $PhoneNumber,
        
        # Street Address
        [Parameter()]
        [String]
        $StreetAddress,
        
        # City
        [Parameter()]
        [string]
        $City,

        # State
        [Parameter()]
        [string]
        $State,

        # Post Code
        [Parameter()]
        [string]
        $PostalCode,

        # Country
        [Parameter()]
        [string]
        $Country,

        # Phone number (Optional)
        [Parameter()]
        [string]
        $Webpage,

        # User to mirror permissions from
        [Parameter(
            Mandatory=$true
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]
        $MirrorUser,

        # The OU to place the user in
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $OU,

        # Password to assign to the user
        [Parameter(
            Mandatory=$true
        )]
        [SecureString]
        $Password
    )
    try 
    {
        New-ADUser `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Name "$($Firstname) $($Lastname)" `
            -DisplayName "$($Firstname) $($Lastname)" `
            -SamAccountName $SamAccountName `
            -UserPrincipalName $UPN `
            -EmailAddress $EmailAddress.ToString() `
            -Description $JobTitle `
            -Title $JobTitle `
            -OfficePhone $PhoneNumber `
            -StreetAddress $StreetAddress `
            -City $City `
            -State $State `
            -PostalCode $PostalCode `
            -Country $Country `
            -HomePage $Webpage `
            -Department $MirrorUser.Department `
            -ScriptPath $MirrorUser.ScriptPath `
            -Path $OU `
            -Enabled $True `
            -AccountPassword $Password `
            -ErrorAction Stop
    }
    catch [UnauthorizedAccessException]
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogString "Could not create the user - please run this script as admin."
        return $null
    }
    catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException]
    {
        Write-NewestErrorMessage -LogType WARNING -CaughtError $_ -LogToFile $false -LogString "Could not assign password to account"        

        $NewUser = Get-NewAccount -SamAccountName $SamAccountName
        if (!$NewUser)
        {
            Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not find new account to assign the password"
            return $null
        }

        Write-Space
        Write-Host "Please provide a new password"
        
        for (;;)
        {
            # Prompt the user for a new password
            $Password = Get-Password

            try 
            {
                Set-ADAccountPassword -Identity $NewUser -Reset -NewPassword $Password
                Enable-ADAccount $NewUser

                Write-Host "`r`n- Successfully set account password. Continuing."
                return $NewUser
            }
            catch
            {
                Write-Space
                Write-Warning "`r`nFailed to set account password. Please try again."
                continue
            }
        }
    }
    catch
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Failed to create new user. Exiting"
        return $null
    }

    # NOTE: Check that this works as expected
    if ($NewUser = Get-NewAccount $SamAccountName)
    {
        return $NewUser
    }
    else
    {
        Write-Space
        Write-Warning "`r`nCould not locate the new account. Please manually check the account before continuing."
        return $null
    }
}

function New-Directory
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # The folder to create the new directory under
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $ParentFolderPath,

        # The Name of the folder to create
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $FolderName
    )
    
    try
    {
        $NewDirectory = New-Item -Path $ParentFolderPath -Name $FolderName -ItemType Directory -ErrorAction Stop
        return $NewDirectory
    }
    catch
    {
        If ($_.CategoryInfo.Category -eq "ResourceExists")
        {
            write-warning "User's folder already exists"

            if (Get-Confirmation "`r`nAn existing folder has been found at $($ParentFolderPath)\$($FolderName)`r`n`r`nAre you sure you want to continue with this folder?`r`n(If you choose 'No', you will need to set up the user's home drive manually)")
            {
                return Get-Item -Path "$($ParentFolderPath)\$($FolderName)"
            }
            else
            {
                return $Null
            }
        }
        else
        {
            Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not create user's folder."
            return $Null
        }  
    }

    return $Null
}

function New-UserFolder
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # Sam name of the user
        [Parameter(
            Mandatory=$true
        )]
        [String]
        $SamAccountName,

        # Domain the user is created under
        [Parameter(
            Mandatory=$true
        )]
        [String]
        $Domain,

        # The folder to create the new directory under
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $ParentFolderPath,

        # The Name of the folder to create
        [Parameter()]
        [string]
        $FolderName = $SamAccountName
    )

    if (!($NewDirectory = New-Directory -ParentFolderPath $ParentFolderPath -FolderName $FolderName))
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not create user's folder."
        return $null        
    }

    if (!(Set-FolderPermissions -SamAccountName $SAMAccountName -Domain $Domain -Path $NewDirectory.FullName))
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not set permissions on the users's folder."
        return $null
    }

    return $NewDirectory
}

function New-HomeDrive
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # The user's Sam Account Name
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $SamAccountName,

        # The domain the user is created under
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Domain,

        # The path that your want the folder created under EXCLUDING the folder name
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $HomeDriveDirectory, 

        # The name of the folder to be created
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $FolderName,
        
        # The drive letter to assign to the Home Drive.
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $DriveLetter
    )

    $HomeDrive = New-UserFolder `
        -SamAccountName $SamAccountName `
        -Domain $Domain `
        -ParentFolderPath $HomeDriveDirectory `
        -FolderName $FolderName

    if ($HomeDrive)
    {
        Set-HomeDrive `
            -Identity $SamAccountName `
            -HomeDrivePath $HomeDrive.Fullname `
            -DriveLetter $DriveLetter
    }
    else 
    {
        Write-Space
        Write-Warning "`r`nFailed to create user's Home Drive. Please check manually."
    }
    
} 

function Set-MobileNumber
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Identity,

        [Parameter()]
        [string]
        $MobileNumber
    )

    if (!$MobileNumber) { return }

    try 
    {
        Set-ADUser -Identity $Identity -MobilePhone $MobileNumber
        Write-Host "- Set Mobile Number field"
    }
    catch 
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not set user's mobile number."
    }
}

function Set-Company
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param(
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Identity,
        
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Company
    )
    
    try
    {
        Set-ADUser -Identity $Identity -Company $Company
        Write-Host "- Set Company field"
    }
    catch
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not set user's Company attribute."
    }
}
<#
function Set-OfficeAddress 
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param ( 
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Identity
    )
    
}
#>
function Set-FolderPermissions
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([System.Boolean])]
    param (
        # Path to the folder
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Path,

        # Sam name of the user
        [Parameter(
            Mandatory=$true
        )]
        [String]
        $SamAccountName,

        # Domain the user is created under
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Domain
    )

    try
    {
        # Get the current permissions of the home drive folder
        $ACL = Get-Acl $Path

        # The new NTFS permissions rule parameters 
        $RuleParameters = @(
            "$($Domain)\$($SamAccountName)",
            "FullControl",
            @(
                "ContainerInherit"
                "ObjectInherit"
            ),
            "None",
            "Allow"
            )

        # Add the rule to the current permissions list
        $Rule = New-Object `
            -TypeName System.Security.AccessControl.FileSystemAccessRule `
            -ArgumentList $RuleParameters

        $ACL.SetAccessRule($Rule) 

        # Set the NTFS permissions on the user's home folder to our new list
        Set-Acl -Path $Path -AclObject $ACL
        
        return $true
    }
    catch
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not set permissions on folder."
        return $false
    }
}

# Copies properties from the mirrored user to the provided account
function Set-MirroredProperties
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # User to set properties on
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Identity,

        # User to mirror properties from
        [Parameter(
            Mandatory=$true
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]
        $MirrorUser
    )
    try
    {
        Set-ADUser -Identity $Identity -Manager $MirrorUser.manager -State $MirrorUser.st -Country $MirrorUser.c -PostalCode $MirrorUser.postalCode -StreetAddress $MirrorUser.streetAddress -City $MirrorUser.l -Office $MirrorUser.physicalDeliveryOfficeName -HomePage $MirrorUser.HomePage        
        Write-Output "- Address and manager have been set:`r`n`r`nManager:`r`n$($MirrorUser.manager)`r`n`r`nAddress:`r`n$($MirrorUser.streetAddress)`r`n$($MirrorUser.l)`r`n$($MirrorUser.st) $($MirrorUser.postalCode)`r`n"
    }
    catch
    {
        Write-NewestErrorMessage -LogType WARNING -CaughtError $_ -LogToFile $true -LogString "Could not set some parameters - please double check the new account's address, and manager"    
    }
}

# Copies groups from the mirrored user to the provided account
function Set-MirroredGroups 
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # User to add to the groups
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Identity,

        # User to mirror groups from
        [Parameter(
            Mandatory=$true
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]
        $MirrorUser,

        # Array of groups to exclude in the mirror
        [Parameter(
            Mandatory=$false
        )]
        [string[]]
        $ExcludedGroups = @(),

        # Array of additional group to add users to. Must be in the form of an object GUID.
        [Parameter(
            Mandatory=$false
        )]
        [string[]]
        $AdditionalGroups = @()
    )

    try 
    {
        $Groups = Get-ADPrincipalGroupMembership $MirrorUser | Where {$_.name -ne 'Domain Users' -and $_.objectGuid -notin $ExcludedGroups}
    }
    catch 
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not get $($MirrorUser.DisplayName)'s groups - please add memberships manually"
        return
    }

    if ($AdditionalGroups)
    {
        $add_groups = foreach ($groupID in $AdditionalGroups) { 
            try 
            {
                get-ADGroup $groupID
            }
            catch 
            {
                Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not locate additional group with object Guid `'$($groupID)`'"
            }
        }

        if ($add_groups) 
        {
            $Groups += $add_groups
        }
    }

    if (!$Groups)
    {
        Write-Space
        Write-Warning "Mirror user `"$($MirrorUser.Name)`" is not part of any groups. Continuing."
        return
    }

    Write-Host "`r`n----------`r`nThe following group memberships will be applied:`r`n----------`r`n"
    foreach ($Group in $Groups)
    {
        Write-Host "- $($Group.Name)"
    }
    Write-Host "----------`r`n"

    #Add account to mirrored user's group memberships
    try
    {
        Add-AdPrincipalGroupMembership -Identity $Identity -MemberOf $Groups
        Write-Output "- Finished adding user's groups"
    }
    catch
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not add user to all groups - please doublecheck group memberships"
    }    
}

function Set-ProxyAddresses
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # The new user to add the addresses to
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Identity,

        # Array of Proxy Addresses       
        [Parameter(
            Mandatory=$true
        )]
        [string[]]
        $ProxyAddresses
    )

    try 
    {
        Set-ADUser -Identity $Identity -Add @{ProxyAddresses = $ProxyAddresses}
        Write-Host "- Set Proxy Addresses"
    }
    catch 
    {
        Write-NewestErrorMessage -LogType WARNING -CaughtError $_ -LogToFile $true -LogString "Could not set Proxy Addresses"
        
    }
}

function Set-LDAPMail
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # The new user to add the field to
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Identity,

        # Primary SMTP Address
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $PrimarySmtpAddress
    )

    try 
    {
        Set-ADUser -Identity $Identity -add @{Mail = $PrimarySmtpAddress}
        Write-Host "- Set Mail field"            
    }
    catch 
    {
        Write-NewestErrorMessage -LogType WARNING -CaughtError $_ -LogToFile $true -LogString "Could not set Mail field"
    }
    
}

function Set-LDAPMailNickName
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # The new user to add the field to
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Identity,

        # Sam Account name
        [Parameter(
            Mandatory=$true
        )]
        [String]
        $SamAccountName
    )

    try 
    {
        Set-ADUser -Identity $Identity -add @{MailNickName = $SamAccountName}
        Write-Host "- Set Mail Nickname field"
    }
    catch 
    {
        Write-NewestErrorMessage -LogType WARNING -CaughtError $_ -LogToFile $true -LogString "Could not set Mail Nickname field"
    }
}

function Set-UserFolderPermissions 
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # Identity of the user to provide access to
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $SamAccountName,

        # Domain the user is created under
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Domain,

        # Path to the folder
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $FolderPath
    )

    # Get the current permissions of the home drive folder
    $ACL = Get-Acl $FolderPath

    # The new NTFS permissions rule parameters 
    $RuleParameters = @(
        "$($Domain)\$($SamAccountName)",
        "FullControl",
        @(
            "ContainerInherit"
            "ObjectInherit"
        ),
        "None",
        "Allow"
        )

    # Add the rule to the current permissions list
    $Rule = New-Object `
        -TypeName System.Security.AccessControl.FileSystemAccessRule `
        -ArgumentList $RuleParameters

    $ACL.SetAccessRule($Rule) 

    # Set the NTFS permissions on the user's home folder to our new list
    Set-Acl -Path $FolderPath -AclObject $ACL
    
}

function Set-HomeDrive 
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # User identity to set the home drive on
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Identity,

        # Path to the folder to set as the home drive
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $HomeDrivePath, 

        # Drive letter to assign (Without colon, or backslash)
        [Parameter()]
        [string]
        $DriveLetter = "H"
    )

    $DriveLetter = $DriveLetter[0] + ":"
        
    try
    {
        # Set home drive on the user's profile
        Set-ADUser -Identity $Identity -HomeDrive $DriveLetter -HomeDirectory $HomeDrivePath -ErrorAction Stop
        Write-Host "`r`n- Successfully set home drive"
    }
    catch
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not set the user profile's home drive"
    }    
}

<#
.SYNOPSIS
    Sets the default retention policy of the given user's mailbox
.DESCRIPTION
    Sets the default retention policy of the given user's mailbox
.EXAMPLE

.INPUTS
    
.OUTPUTS

#>
function Set-MailboxDefaultRetentionPolicy 
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        # The identity of the user mailbox
        [Parameter(
            Mandatory=$true            
        )]
        [string]
        $Identity,

        # The name of the retention policy to set as default
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $PolicyName
    )
    
    try 
    {
        $null = Set-Mailbox -Identity $Identity -RetentionPolicy $PolicyName -ErrorAction Stop
    }
    catch 
    {
        Write-NewestErrorMessage -LogType Warning -LogToFile $true -CaughtError $_ -LogString "Could not set the default retention policy"
    }
}

<#
.SYNOPSIS
    Adds the given user to the given groups
.DESCRIPTION
    Checks the groups provided, and adds the user if they are missing from any of them
.EXAMPLE
    PS C:\> Add-GroupMemberships -Identity "JDoe" -Groups "Group_1", "Group-2"
    Explanation of what the example does
#>
function Add-GroupMemberships
{
    param (
        # Identy of the user to add
        [Parameter(
            Mandatory=$true,
            Position=0
        )]
        [string]
        $Identity,

        # Array of groups to add the user to
        [Parameter(
            Mandatory=$true,
            Position=1
        )]
        [string[]]
        $Groups
    )

    # Do not continue if there are no groups
    if ($Groups.Count -lt 1) { return }
    
    $memberships = Get-ADPrincipalGroupMembership $Identity | select Name, SamAccountName        

    # Check for existing group memberships, and add them if they are missing
    foreach ($group in $Groups) 
    {
        if ($memberships.SamAccountName -notcontains $group)
        {
            try 
            {
                Add-ADPrincipalGroupMembership -Identity $Identity -MemberOf $group
            }                
            catch 
            { 
                Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Something went wrong while adding the user to group: $(group)"
            }
        }
    }    
}

<#
.SYNOPSIS
    Returns the name of the Mailbox Database with the greatest amount of free space available from the values provided. Returns null if nothing could be found.
.DESCRIPTION
    Returns the name of the Mailbox Database with the greatest amount of free space available from the values provided. Returns null if nothing could be found.
.EXAMPLE
    
.INPUTS
    
.OUTPUTS
    
#>
function Get-MostAvailableMailboxDatabase 
{
    param (
        # The names of the databases to compare
        [Parameter(
            Mandatory=$true
        )]
        [string[]]
        $DatabaseNames       
    )
    
    $MostAvailable = $null

    foreach ($name in $DatabaseNames)
    {
        try 
        {
            $database = Get-MailboxDatabase -Identity $name -Status | select name, AvailableNewMailboxSpace
            
            if ($null -eq $MostAvailable)
            {
                $MostAvailable = $database
            }
            elseif ($MostAvailable.AvailableNewMailboxSpace -lt $database.AvailableNewMailboxSpace)
            {
                $MostAvailable = $database
            }
        }
        catch
        {
            Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Something went wrong while getting the most available mailbox database"
        }
    }

    #Remove-ExchangeSnapins
    
    if ($null -ne $MostAvailable)
    { 
        return $MostAvailable.name
    }
    else 
    {
        return $null
    }    
}

<#
.SYNOPSIS
    Creates/Enables an existing user's mailbox
.DESCRIPTION
    Creates/enables an existing user's mailbox. 
    User name paramaters can optionally be passed in to avoid Exchange using dirty name values to create the user's addresses
.EXAMPLE
    PS C:\> Enable-UserMailbox -Firstname "Test" -Lastname "Mc'Account" -FirstnameClean "Test" -LastnameClean "McAccount" -Alias "Test.McAccount" -Database "YOUREXCHANGEDATABASE" -ExchangeYear "2010"

    Note: The user's name parameters are the same as the output of Get-Fullname, and can be passed via splatting for convenience. 
    Eg.

    PS C:\> $Name = Get-Fullname
    PS C:\> Enable-UserMailbox @Name -Alias "Test.McAccount" -Database "YOUREXCHANGEDATABASE" -ExchangeYear "2010"
.INPUTS
    String
.PARAMETER Identity
    The identity of the user whose' mailbox should be enabled.
.PARAMETER Firstname
    Raw first name of the user. This can be a raw form of the user's name for display purposes. The user's name will be set to this at the end of this function.
.PARAMETER Lastname
    Raw last name of the user. This can be a raw form of the user's name for display purposes. The user's name will be set to this at the end of this function.
.PARAMETER FirstnameClean
    Optional. Cleaned first name - this may be used by exchange to set up the user's email address, so it should be free of any illegal characters/symbols.
.PARAMETER LastnameClean
    Optional. Cleaned last name - this may be used by exchange to set up the user's email address, so it should be free of any illegal characters/symbols.
.PARAMETER Alias
    The alias of the user (The name before the @ symbol in their address).
.PARAMETER Database
    The database the user is to be enabled on.
.PARAMETER ExchangeYear
    The year version of the exchange. Current valid options are "2010" or "2013"
#>
function Enable-UserMailbox
{
    # Default parameter set to use
    [CmdletBinding(DefaultParametersetName="NoName")]
    param (
        # The UPN of the user the mailbox belongs to
        [Parameter(
            Mandatory=$true,
            ParameterSetName="Name"
        )]
        [Parameter(
            Mandatory=$true,
            ParameterSetName="NoName",
            Position=0
        )]
        [string]
        $Identity,

        # Raw first name of the user (what the user's name/display name will be when finished)
        [Parameter(
            Mandatory=$true,
            ParameterSetName="Name"
        )]
        [string]
        $Firstname,

        # Raw last name of the user (what the user's name/display name will be when finished)
        [Parameter(
            Mandatory=$true,
            ParameterSetName="Name"
        )]
        [string]
        $Lastname,

        # Clean first name of the user (illegal characters removed) - used while setting up the user's mailbox/email address
        [Parameter(
            Mandatory=$true,
            ParameterSetName="Name"
        )]
        [string]
        $FirstnameClean,

        # Clean last name of the user (illegal characters removed) - used while setting up the user's mailbox/email address
        [Parameter(
            Mandatory=$true,
            ParameterSetName="Name"
        )]
        [string]
        $LastnameClean,

        # The user's mail name (before the @ symbol)
        [Parameter(
            Mandatory=$true,
            ParameterSetName="Name"
        )]
        [Parameter(
            Mandatory=$true,
            ParameterSetName="NoName"
        )]
        [String]
        $Alias, 

        # The exchange database to use
        [Parameter(
            Mandatory=$true,
            ParameterSetName="Name"
        )]
        [Parameter(
            Mandatory=$true,
            ParameterSetName="NoName"
        )]
        [String]
        $Database
    )

    [bool] $SetName = $PSCmdlet.ParameterSetName -eq "Name"

    if ($SetName)
    {
        try 
        {
            Set-ADUser -Identity $Identity -GivenName $FirstnameClean -Surname $LastnameClean -DisplayName "$($FirstnameClean) $($LastnameClean)"
        }
        catch 
        {
            Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not set user's name before enabling mailbox.`r`nPlease double check the proxy addresses."
        }
    }

    try 
    {
        $MailboxResult = Enable-Mailbox -Identity $Identity -Alias $Alias -Database $Database -ErrorAction Stop
        if ($MailboxResult)
        {
            Write-Host "`r`n- Successfully enabled user mailbox."
        }        
    }
    catch
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not enable mailbox."
    }

    if ($SetName)
    {
        try 
        {
            Set-ADUser -Identity $Identity -GivenName $Firstname -Surname $Lastname -DisplayName "$($Firstname) $($Lastname)"
        }
        catch 
        {            
            Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not set user's name back after enabling mailbox.`r`nPlease double check user's name."
        }
    }
}

<#
.SYNOPSIS
    Enables an in-place archive for the specified user mailbox, on the specified database
.DESCRIPTION
    Enables an in-place archive for the specified user mailbox, on the specified database
.EXAMPLE
    
.INPUTS
    
.OUTPUTS

#>
function Enable-MailboxArchive 
{
    param (
        # Identity of the user/mailbox
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Identity,

        # Database to enable to archive on
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $Database
    )

    try 
    {
        $null = Enable-mailbox $Identity -archive -ArchiveDatabase $Database -ErrorAction Stop
    }
    catch 
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "A problem occurred while trying to enable the mailbox archive"
    }    
}

function Remove-ExchangeSnapins 
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    foreach ($Snapin in (Get-PSSnapin))
    {
        if ($Snapin.Name -like "*Exchange.Management*")
        {
            try 
            {
                Remove-PSSnapin $Snapin.Name                
            }
            catch 
            {
                Write-NewestErrorMessage -LogType WARNING -CaughtError $_ -LogToFile $true -LogString "Could not remove Exchange Snapin `"$($Snapin.name)`""
            }
        }
    }    
}
    
function Start-O365Sync 
{
    Start-ADSyncSyncCycle -PolicyType Delta    
}