$ErrorsPresent = $false
$LogPath = 'C:\temp\UserCreationADLog.txt'

Add-Type -TypeDefinition @"
   public enum LogTypes
   {
      DEBUG,
      WARNING,
      ERROR
   }
"@

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
            {'no','n' -icontains $_}{return $false}
            default
            {
                Write-Warning "I could not understand your response. Please reply with either yes or no."
                continue
            }            
        }
    }
}

function Write-NewestErrorMessage
{
    param(
        [Parameter(Mandatory=$true)]
        [LogTypes] $LogType, 
            
        [Parameter()]
        [String] $LogString    
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
    Write-Host -ForegroundColor $LogColour "$($LogType.ToString()): $($Error[0].Exception.Message)`r`n"
    
    $LogFileText = "$(Get-Date -Format "yyyy/MM/dd | HH:mm:ss") | $($LogType.ToString()) | $($LogString) : $($Error[0].Exception.Message)"
    Out-File -FilePath $LogPath -Append -InputObject $LogFileText

    if ($LogType.ToString() -ne "DEBUG")
    {
        Write-Warning "`r`nA log has been generated and can be found at $($LogPath).`r`nIf this was unexpected, please send this log to the maintainer."
    }
}

# Import AD Module
try
{
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch
{
    Write-NewestErrorMessage -LogType ERROR -LogString "Could not import Active Directory Module. Aborting."
    return
}
<#
# Check if we are running as admin
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (!$currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    Write-Warning "No admin priviledges detected.`r`n"
    if (!(Get-Confirmation "WARNING: This script is not running as admin, you will probably be unable to create the user account.`r`n`r`nAre you sure you want to continue?"))
    {
        return
    }    
}
#>
function Write-Space
{
    Write-Host ""
}

function Get-FirstName
{
    [string]$FirstName = Read-Host "----------`r`nPlease enter a First Name"
    return $FirstName
}

function Get-Lastname
{
    [string]$LastName = Read-Host "----------`r`nPlease enter a Last Name"
    return $LastName
}

Function Get-JobTitle
{
    do 
    {
        $JobTitle = Read-Host "----------`r`nPlease enter a Job Description"
    } while (!(Get-Confirmation "Is this correct?`r`n`r`n$($JobTitle)"))
    return $JobTitle
}

function Get-PhoneNumber
{
    do 
    {
        $PhoneNumber = Read-Host "----------`r`nPlease enter a phone number (Leave blank for none)"        
    } while (($PhoneNumber -ne "") -and ($Null -ne $PhoneNumber) -and !(Get-Confirmation "Is this correct?`r`n`r`n$($PhoneNumber)"))
    return $PhoneNumber
}

function Get-Password
{
    param (
        # The password to set it to
        [Parameter(Position=0)]
        [string]
        $Password = '$Password99!'
    )
    
    # Allow for a custom password
    if (!(Get-Confirmation "----------`r`nWould you like to set the default password?`r`n(Choose 'No' to create your own)`r`n`r`n$($Password)"))
    {
        do
        {
            $Password = Read-Host "Please enter a Password"
        } while ($Password -eq "" -or $null -eq $Password)
    }
    return $Password
}

# Gets the username 
function Get-MirrorUser 
{
    param (
        # The format of the username as a text prompt
        [Parameter(
            Mandatory=$true)]
        [string]
        $UsernameFormat
    )

    for (;;)
    {        
        $MirrorName = Read-Host "`r`n----------`r`nPlease enter a user to mirror `r`n(Username format: $($UsernameFormat))"
        $MirrorName = $MirrorName.Trim()
        
        try 
        {
            $User = Get-ADUser -Identity $MirrorName -Properties *

            Write-Host "`r`nLocated mirror account:`r`n`r`nAccount`r`n$($User.SAMAccountName)`r`n`r`nName`r`n$($User.Name)`r`n`r`nEmail`r`n$($User.EmailAddress)"

            if (Get-Confirmation "Is this the correct account?")
            {
                break
            }            
        }
        catch 
        {
            Write-Host ""
            Write-NewestErrorMessage -LogType WARNING -LogString "Could not find any user with that username."
            continue
        }
    }

    return $User
}

function Get-OU 
{
    param (
        # User to get the OU from
        [Parameter(
            Mandatory=$true
        )]
        [string]
        $MirrorUser
    )
    
    return ($MirrorUser.DistinguishedName -replace '^cn=.+?(?<!\\),')
}

# Optimises the given name by removing leading/trailing white space, illegal characters, and capitalising the first letter if required
function Optimize-Name
{
    param (
        # Name to optimise
        [Parameter(
            Mandatory=$true,
            Position=0)]
        [string]
        $Name,

        # Whether to capitalise the name
        [Parameter()]
        [bool]
        $FixCapitalisation = $false
    )
    
    # illegal characters
    $CharList = "'", "`"", "/", "\",";", ":", "(", ")", "[", "]", "!", "@", "$", "%", "^", "&", "*", "``", "~", "."
    
    # Remove illegal characters
    foreach ($Char in $CharList)
    {
        $Name = $Name.Replace($Char, "")
    }

    #Remove WhiteWrite-Spaces
    $Name = $Name.Trim()

    if ($FixCapitalisation)
    {
        $Name = (Get-Culture).TextInfo.ToTitleCase($Name.ToLower())
    }

    return $Name
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
        get-ADUser -identity $SAM > $null
        return $false
    } catch 
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

# Confirms all new user details
function Confirm-NewUserDetails 
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

        [Parameter(
            Mandatory=$true)]
        [string]
        $Password,

        [Parameter(
            Mandatory=$true)]
        [string]
        $MirrorUser

    )

    $ConfirmationMessage = "You are about to create an account with the following details:`r`n`r`nFirst Name: $($Firstname)`r`n`r`nLast Name: $($Lastname)`r`n`r`nJob Title: $($JobTitle)`r`n`r`nUsername: $($SamAccountName)`r`n`r`nEmail Address: $($EmailAddress)`r`n`r`nAccount Password: $($Password)`r`n`r`nAccount to Mirror: $($MirrorUser.DisplayName)`r`n`r`nDo you wish to proceed?"
    return (Get-Confirmation $ConfirmationMessage)
}

function Test-NewAccountExists
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

    Write-Host "Waiting for account to become available..."

    for ($i = 0; $i -lt $AttemptCount; $i++)
    {
        try 
        {
            Get-ADUser $SamAccountName > $null
            return $true
        } 
        catch 
        { 
            start-sleep -Seconds 3
            continue
        }
    }

    return $false
}

function New-UserAccount {
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

        # Job title/account description (Optional)
        [Parameter()]
        [string]
        $JobTitle,

        # Phone number (Optional)
        [Parameter()]
        [string]
        $PhoneNumber,

        # User to mirror permissions from
        [Parameter(
            Mandatory=$true
        )]
        [string]
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
        [string]
        $Password
    )
    try 
    {
        New-ADUser -GivenName $Firstname -Surname $Lastname -Name "$($Firstname) $($Lastname)" -DisplayName "$($Firstname) $($Lastname)" -SamAccountName $SamAccountName -UserPrincipalName $UPN -Description $JobTitle -Title $JobTitle -OfficePhone $PhoneNumber -Department $MirrorUser.Department -Path $OU -Enabled $True -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force)
    }
    catch [UnauthorizedAccessException]
    {
        WriteNewestErrorMessage -LogType ERROR -LogString "Could not create the user - please run this script as admin."
        return $false
    }
    catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException]
    {
        $Script:PasswordInvalid = $True

        Write-Warning "`r`nCould not assign password to new user. Please enter a new password."

        if (!(Test-NewAccountExists))
        {
            return $false
        }
        
        for (;;)
        {
            # Get a new password
            $Password = Read-Host "`r`nPlease enter a password"
            
            try 
            {                
                $NewUser = Get-ADUser $SAM
            }
            catch 
            {
                Write-NewestErrorMessage -LogType ERROR -LogString "Could not find new account. Exiting"
                return $false
            }

            try 
            {
                Set-ADAccountPassword -Identity $NewUser -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force)
                Enable-ADAccount $NewUser

                Write-Host "`r`n- Successfully set account password. Continuing."
                break                
            }
            catch
            {
                WriteNewestErrorMessage -LogType WARNING -LogString "Failed to set account password. Please try again."
                continue
            }
        }
    }
    catch
    {    
        WriteNewestErrorMessage -LogType ERROR -LogString "Failed to create new user. Exiting"
        return $false
    }
    
}