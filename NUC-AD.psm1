$ErrorsPresent = $null
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
            {'no','n' -icontains $_}{return $null}
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
    # Import-Module ActiveDirectory -ErrorAction Stop
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

# Writes an empty line to the console
function Write-Space
{
    Write-Host ""
}

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

# Prompt the user to provide a first name for the new account
function Get-FirstName
{
    [string]$FirstName = Read-Host "----------`r`nPlease enter a First Name"
    return $FirstName
}

# Prompt the user to provide a last name for the new account
function Get-Lastname
{
    [string]$LastName = Read-Host "----------`r`nPlease enter a Last Name"
    return $LastName
}

# Get the first and last name, and loop until it is approved. Returns an array containing the First and Last name
function Get-FullName
{
    param (
        # Choose whether to prompt for, and fix capitalisation
        [Parameter()]
        [ParameterType]
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
            $ShouldFixCapitalisation = Get-Confirmation "----------`r`nWould you like to fix/standardise the capitalisation of this name?`r`n(Choose yes if you are not sure)"    
        }
        else 
        {
            $ShouldFixCapitalisation = $null
        }

        # Clean the given names
        $Firstname = Optimize-Name -Name $Firstname -FixCapitalisation $ShouldFixCapitalisation
        $Lastname = Optimize-Name -Name $LastName -FixCapitalisation $ShouldFixCapitalisation

        # Confirm the name is correct
        If (Confirm-Name -FirstName $Firstname -LastName $Lastname)
        {
            return @($Firstname, $Lastname)
        }
    }
}

# Prompt the user to provide a job title/description for the new account
Function Get-JobTitle
{
    do 
    {
        $JobTitle = Read-Host "----------`r`nPlease enter a Job Description"
    } while (!(Get-Confirmation "Is this correct?`r`n`r`n$($JobTitle)"))
    return $JobTitle
}

# Prompt the user to provide a phone number for the new account
function Get-PhoneNumber
{
    do 
    {
        $PhoneNumber = Read-Host "----------`r`nPlease enter a phone number (Leave blank for none)"        
    } while (($PhoneNumber -ne "") -and ($Null -ne $PhoneNumber) -and !(Get-Confirmation "Is this correct?`r`n`r`n$($PhoneNumber)"))
    return $PhoneNumber
}

# Prompt the user to provide a password for the new account
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
        $PrimaryAddress = "SMTP:$($Mail)@$($PrimaryDomain)"
        $ProxyAddresses = @($PrimaryAddress)
        
        foreach ($Domain in $SecondaryDomains)
        {
            $ProxyAddresses += "smtp:$($Mail)@$($Domain)"            
        }

        if (Confirm-PrimarySMTPAddress -PrimarySMTP $EmailAddress) 
        {
            return $ProxyAddresses
        }
        else 
        {
            $Script:Mail = Read-Host "Enter the email address (Only enter the text BEFORE the '@' sign)"    
        }
    }
}

# Gets the username 
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
                break
            }            
        }
        catch 
        {
            Write-Space
            Write-NewestErrorMessage -LogType WARNING -LogString "Could not find any user with that username."
            continue
        }
    }

    return $User
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
            $User = Get-ADUser $SamAccountName
            break
        } 
        catch 
        { 
            start-sleep -Seconds 3
            continue
        }
    }

    if ($User)
    {        
        return $User
    }
    else
    {
        return $null
    }
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
        $FixCapitalisation = $null
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
        return $null
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

# Creates a new user from the parameters provided. Returns true if the account was created successfully, and false if not.
function New-UserAccount 
{
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
        New-ADUser -GivenName $Firstname -Surname $Lastname -Name "$($Firstname) $($Lastname)" -DisplayName "$($Firstname) $($Lastname)" -SamAccountName $SamAccountName -UserPrincipalName $UPN -Description $JobTitle -Title $JobTitle -OfficePhone $PhoneNumber -Department $MirrorUser.Department -Path $OU -Enabled $True -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) -ErrorAction Stop
    }
    catch [UnauthorizedAccessException]
    {
        Write-NewestErrorMessage -LogType ERROR -LogString "Could not create the user - please run this script as admin."
        return $null
    }
    catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException]
    {
        Write-Warning "`r`nCould not assign password to new user."

        $NewUser = Get-NewAccount -SamAccountName $SAM
        if (!$NewUser)
        {
            Write-NewestErrorMessage -LogType ERROR -LogString "Could not find new account."
            return $null
        }
        
        for (;;)
        {
            # Prompt the user for a new password
            $Password = Read-Host "`r`nPlease enter a password"

            try 
            {
                Set-ADAccountPassword -Identity $NewUser -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force)
                Enable-ADAccount $NewUser

                Write-Host "`r`n- Successfully set account password. Continuing."
                return $NewUser
            }
            catch
            {
                Write-NewestErrorMessage -LogType WARNING -LogString "Failed to set account password. Please try again."
                continue
            }
        }
    }
    catch
    {    
        Write-NewestErrorMessage -LogType ERROR -LogString "Failed to create new user. Exiting"
        return $null
    }

    # NOTE: Check that this works as expected
    if ($NewUser = Get-NewAccount $SAM)
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

# Copies properties from the mirrored user to the provided account
function Set-MirroredProperties
{
    param (
        # User to set properties on
        [Parameter(
            Mandatory=$true
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]
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
        Set-ADUser -Identity $NewUser.identity -Manager $MirrorUser.manager -State $MirrorUser.st -Country $MirrorUser.c -PostalCode $MirrorUser.postalCode -StreetAddress $MirrorUser.streetAddress -City $MirrorUser.l -Office $MirrorUser.physicalDeliveryOfficeName -HomePage $MirrorUser.HomePage
        Write-Output "- Address and manager have been set."
    }
    catch
    {
        WriteNewestErrorMessage -LogType WARNING -LogString "Could not set some parameters - please double check the account's address and manager"    
    }
}

# Copies groups from the mirrored user to the provided account
function Set-MirroredGroups 
{
    param (
        # User to add to the groups
        [Parameter(
            Mandatory=$true
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]
        $Identity,

        # User to mirror groups from
        [Parameter(
            Mandatory=$true
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]
        $MirrorUser
    )
    #Add account to mirrored user's group memberships
    try
    {
        $NewUser | Add-AdPrincipalGroupMembership -MemberOf (Get-ADPrincipalGroupMembership $MirrorUser | Where {$_.name -ne 'Domain Users'})
        Write-Output "- User has been added to all mirrored user's groups"
    }
    catch
    {
        WriteNewestErrorMessage -LogType ERROR -LogString "Could not add user to all groups - please doublecheck group memebrships"
    }
    
}

function Set-ProxyAddresses
{
    param (
        # The new user to add the addresses to
        [Parameter(
            Mandatory=$true
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]
        $Identity,

        # Array of Proxy Addresses       
        [Parameter(
            Mandatory=$true
        )]
        [string[]]
        $ProxyAddresses,
    )

    try 
    {
        Set-ADUser -Identity $NewUser -Add @{ProxyAddresses = $ProxyAddresses}
        Write-Host "- Set Proxy Addresses"
    }
    catch 
    {
        Write-Warning "Could not set Proxy Addresses"
        
    }
}

function Set-LDAPMail
{
    param (
        # The new user to add the field to
        [Parameter(
            Mandatory=$true
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]
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
        Set-ADUser -Identity $NewUser -add @{Mail = $PrimarySmtpAddress}
        Write-Host "- Set Mail field"            
    }
    catch 
    {
        Write-Warning "Could not set Mail field"
    }
    
}

function Set-LDAPMailNickName
{
    param (
        # The new user to add the field to
        [Parameter(
            Mandatory=$true
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]
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
        Set-ADUser -Identity $NewUser -add @{MailNickName = $SAM}
        Write-Host "- Set Mail Nickname field"
    }
    catch 
    {
        Write-Warning "Could not set Mail Nickname field"
    }
}
    
function Start-O365Sync 
{
    Start-ADSyncSyncCycle -PolicyType Delta    
}