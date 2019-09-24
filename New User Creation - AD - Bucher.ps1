Write-Output "Loading - please wait`r`n"

<#
----------
# Variables 
----------
#>

$ErrorsPresent = $false
$LogPath = 'C:\temp\UserCreationADLog.txt'

<#
----------
# Data Types
----------
#>

Add-Type -TypeDefinition @"
   public enum LogTypes
   {
      DEBUG,
      WARNING,
      ERROR
   }
"@

<#
----------
# Functions
----------
#>

# Client-specific account conventions
# NOTE: Edit these to change things like logon name, and email address.

# DO NOT CHANGE THESE ELSEWHERE IN THE SCRIPT. Do it here instead
function SetClientSpecificParams()
{
    param(
        [Parameter(Mandatory=$true)]
        [string] $FirstName,
            
        [Parameter()]
        [String] $LastName
    ) 

    #Remove illegal characters
    $CharList = "'", "`"", "/", "\",";", ":", "(", ")", "[", "]", "!", "@", "$", "%", "^", "&", "*", "``", "~", "."

    foreach ($Char in $CharList)
    {
        $FirstName = $FirstName.Replace($Char, "")
        $LastName = $LastName.Replace($Char, "")
    }

    # Change these to suit the client

    # The format of the username. 
    # This is part of a text prompt for the mirrored user's account (Same format as the SAM parameter)
    [string]$Script:UsernameFormat = "Firstname.Lastname"
    
    # Email 
    # (Only used below in primary/alias SMTP addresses, this can change later in the script)
    # LDAP Fields:   
    # - N/A 
    $Script:Mail = $Firstname+"."+$LastName

    # User logon name/User Principle Name 
    # WARNING: This is different to $Mail
    # LDAP Fields:
    # - UserPrincipleName (User Logon Name)
    $Script:UPN = "$($FirstName).$($LastName)@Macdonald-Johnston.com.au"
    
    # Pre-windows 2000 Logon name
    # LDAP Fields:
    # - SamAccountName    
    $Script:SAM = $FirstName+"."+$LastName

    GenerateProxyAddresses
}

# Prepares email address variables
function GenerateProxyAddresses()
{            
    # Used as the primary SMTP address (this is set with the proxy addresses below)
    # - Mail (EmailAddress)
    # - PrimarySmtpAddress (not explicitly set, but by extension of proxy addresses
    $Script:EmailAddress = "$($FirstName).$($LastName)@buchermunicipal.com.au"

    # Array of mail aliases
    $Script:ProxyAddresses = "SMTP:$($EmailAddress)", "smtp:$($Mail)@Bucher.com.au", "smtp:$($Mail)@JDMacdonald.com.au", "smtp:$($Mail)@MacdonaldJohnston.com.au", "smtp:$($Mail)@Macdonald-Johnston.com.au", "smtp:$($Mail)@MJE.com.au"
}

# Prompts for a yes or no choice. Returns True or False for Yes or No
function PromptForConfirmation([string]$Message)
{
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

# Writes the provided message, plus details of the last error to console and file.
function WriteNewestErrorMessage
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
    Write-Output "`r`nFull Details:"
    Write-Host -ForegroundColor $LogColour "$($LogType.ToString()): $($Error[0].Exception.Message)`r`n"
    
    $LogFileText = "$(Get-Date -Format "yyyy/MM/dd | HH:mm:ss") | $($LogType.ToString()) | $($LogString) : $($Error[0].Exception.Message)"
    Out-File -FilePath $LogPath -Append -InputObject $LogFileText

    if ($LogType.ToString() -ne "DEBUG")
    {
        Write-Warning "`r`nA log has been generated and can be found at $($LogPath).`r`nIf this was unexpected, please send this log to the maintainer."
    }
}

<#
----------
# Pre-run checks
----------
#>

try
{
    add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction Stop
}
catch
{
    if ($_.Exception -like "*Microsoft.Exchange.Management.PowerShell.E2010 because it is already added*") { }
    else
    {
        WriteNewestErrorMessage -LogType ERROR -LogString "Could not import Exchange Management module."
        if (!(PromptForConfirmation "Could not import the Exchange Management module. This script will be unable to enable the user's mailbox. Do you want to continue?"))
        {
            return
        }
    }
}

try
{
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch
{
    WriteNewestErrorMessage -LogType ERROR -LogString "Could not import Active Directory Module. Aborting."
    return
}

# Check if we are running as admin
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (!$currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    Write-Warning "No admin priviledges detected.`r`n"
    if (!(PromptForConfirmation "WARNING: This script is not running as admin, you will probably be unable to create the user account.`r`n`r`nAre you sure you want to continue?"))
    {
        return
    }
    
}

Write-Warning "It is not recommended to create service accounts with this tool."

<#
----------
# Get new user details
----------
#>

Write-Output "`r`n----------`r`nStarting User Creation Process`r`n----------`r`n"

do 
{
    $ReadyToCreate = $false

    # Get the user's name
    do
    {

        [string]$FirstName = Read-Host "----------`r`nPlease enter a First Name"
        Write-Output ""
        [string]$LastName = Read-Host "Please enter a Last Name"

        #Remove whitespaces
        $Firstname = $Firstname.Trim()
        $Lastname = $Lastname.Trim()
        
        If(PromptForConfirmation "Would you like to fix/standardise the capitalisation of this name?`r`n(Choose yes if you are not sure)")
        {
            # Capitalise (only) the first letter of the names
            $FirstName = (Get-Culture).TextInfo.ToTitleCase($FirstName.ToLower())
            $LastName = (Get-Culture).TextInfo.ToTitleCase($LastName.ToLower())
        }
    } while (!(PromptForConfirmation "Is this name correct?`r`n`r`n$($FirstName) $($LastName)"))

    Write-Output ""
    
    # Get the Job Title
    do 
    {
        $JobTitle = Read-Host "----------`r`nPlease enter a Job Description"
    } while (!(PromptForConfirmation "Is this correct?`r`n`r`n$($JobTitle)"))
    
    Write-Output ""

    # Get the phone number
    do 
    {
        $PhoneNumber = Read-Host "----------`r`nPlease enter a phone number (Leave blank for none)"        
    } while (($PhoneNumber -ne "") -and !(PromptForConfirmation "Is this correct?`r`n`r`n$($PhoneNumber)"))

    # Spacing
    Write-Output ""

    # Set the default password
    $Password = "`$Password99!"

    # Allow for a custom password
    if (!(PromptForConfirmation "----------`r`nWould you like to set the default password?`r`n(Choose 'No' to create your own)`r`n`r`n$($Password)"))
    {
        do
        {
            $Password = Read-Host "Please enter a Password"
        } while ($Password -eq "" -or $null -eq $Password)
    }

    # Set client specific parameters, such as email, and logon name
    SetClientSpecificParams -FirstName $FirstName -LastName $LastName

    # Check to see if this account already exists
    try 
    {
        get-ADUser -identity $SAM > $null

        # This is a bit janky, don't judge me!
        try
        {
            throw "The user account already exists."
        }
        catch
        {
            WriteNewestErrorMessage -LogType ERROR -LogString "The user account already exists. Exiting"
            return
        }
    } catch {}
        
    while (!(PromptForConfirmation "----------`r`nThe primary SMTP/email address as been set to:`r`n`r`n$($EmailAddress)`r`n`r`nIs this correct?"))
    {        
        $Mail = Read-Host "Enter the email address (Only enter the text BEFORE the '@' sign)"
        GenerateProxyAddresses  
    } 

    # Spacing
    Write-Output "`r`n----------`r`n"

    # Get the user account to mirror
    for(;;)
    {
        $MirrorName = Read-Host "Please enter a user to mirror `r`n(Username format: $($UsernameFormat))"
        $MirrorName = $MirrorName.Trim()

        try 
        {
            $MirrorUser = get-aduser -identity $MirrorName -Properties * -ErrorAction Stop
        }
        catch 
        {
            WriteNewestErrorMessage -LogType ERROR -LogString "Could not get the specified user"
            $MirrorUser = $null
            continue
        }

        Write-Output "`r`n- Located mirror account:`r`n`r`nAccount:"$MirrorUser.SAMAccountName"`r`nName"$MirrorUser.Name"`r`nEmail"$MirrorUser.EmailAddress

        if (PromptForConfirmation "Is this the correct account?")
        {
            break
        }
    } 
    
    $OU = $MirrorUser.DistinguishedName -replace '^cn=.+?(?<!\\),'

    Write-Output "`r`n----------`r`nVerification`r`n----------`r`n"

    # Confirm
    $ConfirmationMessage = "You are about to create an account with the following details:`r`n`r`nFirst Name: $($Firstname)`r`nLast Name: $($Lastname)`r`n`r`nJob Title: $($JobTitle)`r`n`r`nEmail Address: $($EmailAddress)`r`n`r`nAccount Password: $($Password)`r`n`r`nAccount to Mirror: $($MirrorUser.DisplayName)`r`n`r`nDo you wish to proceed?"

    If (PromptForConfirmation($ConfirmationMessage))
    {
        $ReadyToCreate = $true        
    }
    else
    {
        Write-Output "`r`n----------`r`nCancelled user creation - restarting`r`n----------`r`n"
    }

} while (!$ReadyToCreate)

<#
----------
Create the user account
----------
#>

Write-Output "`r`n----------`r`nBeginning Account Creation`r`n----------`r`n"

try 
{
    New-ADUser -GivenName $Firstname -Surname $Lastname -Name $Firstname" "$Lastname -SamAccountName $SAM -Description $JobTitle -DisplayName "$($Firstname) $($Lastname)" -OfficePhone $PhoneNumber -Department $MirrorUser.Department -Path $OU -Title $JobTitle -UserPrincipalName $UPN -Enabled $True -AccountPassword (ConvertTo-SecureString $password -AsPlainText -force)
}
catch [UnauthorizedAccessException]
{
    WriteNewestErrorMessage -LogType ERROR -LogString "Could not create the user - please run this script as admin."
    return
}
catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException]
{
    $Script:PasswordInvalid = $True

    WriteNewestErrorMessage -LogType WARNING -LogString "Could not assign password to new user. Please wait for the account to become available, and enter a new password."
    Write-Output "`r`n- Waiting for account to become available (Max 60 seconds)."

    # Wait for the account to be found
    for ($i = 0; $i -lt 20; $i++)
    {
        try 
        {
            Get-ADUser -Identity $SAM > $null
            $FailedCheck = $false
        }
        catch 
        {
            $FailedCheck = $true
        }

        if ($FailedCheck)
        {
            start-sleep -seconds 3	    
        }
        else
        {
	        break
        }
    }

    # Confirm we found the account
    if ($FailedCheck)
    {
        Write-host -ForegroundColor Red "`r`nERROR:`r`nFailed to find the new user account. Please remove the account manually (If necessary) before running this script again.`r`n`r`nAborting."
        return
    }

    for (;;)
    {
        # Get a new password
        $Password = Read-Host "`r`nPlease enter a password"

        try 
        {
            $NewUser = Get-ADUser $SAM
            Set-ADAccountPassword -Identity $NewUser -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force)
            Enable-ADAccount $NewUser

            Write-Output "`r`n- Successfully set account password. Continuing."
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
    return
}

Start-Sleep -Seconds 3

<#
----------
Populate the new account's properties
----------
#>

# Get the new user account
Write-Output "`r`n- Waiting for account to become available (Max 60 seconds).`r`n"

# Wait for the account to be found
for ($i = 0; $i -lt 20; $i++)
{
    try 
    {
        $NewUser = Get-ADUser -Identity $SAM -Properties *
        $FailedCheck = $false
    }
    catch 
    {
        $FailedCheck = $true
    }

    if ($FailedCheck)
    {
        start-sleep -seconds 3	    
    }
    else
    {
	    break
    }
}

# Confirm we found the account
if ($FailedCheck)
{
    Write-host -ForegroundColor Red "`r`nERROR:`r`nFailed to find the new user account. Please remove the account manually (If necessary) before running this script again.`r`n`r`nAborting."
    return
}

Write-Output "`r`n- Successfully created the new user account."
Write-Output "`r`n----------`r`nPopulating account details.`r`n----------`r`n"

# Copy mirror user's details
try
{
    $NewUser | Set-ADUser -Manager $MirrorUser.manager -State $MirrorUser.st -Country $MirrorUser.c -PostalCode $MirrorUser.postalCode -StreetAddress $MirrorUser.streetAddress -City $MirrorUser.l -Office $MirrorUser.physicalDeliveryOfficeName -HomePage $MirrorUser.HomePage
    Write-Output "- Address and manager have been set."
}
catch
{
    WriteNewestErrorMessage -LogType ERROR -LogString "Could not set some parameters - please double check the account's address and manager"
    $ErrorsPresent = $True
}

#Add account to mirrored user's group memberships
try
{
    $NewUser | Add-AdPrincipalGroupMembership -MemberOf (Get-ADPrincipalGroupMembership $MirrorUser | Where {$_.name -ne 'Domain Users'})
    Write-Output "- User has been added to all mirrored user's groups"
}
catch
{
    WriteNewestErrorMessage -LogType ERROR -LogString "Could not add user to all groups - please doublecheck group memebrships"
    $ErrorsPresent = $True
}

<#
----------
Set the Home Drive
----------
#>

Write-Output "`r`n`----------`r`nCreating user's home drive`r`n`----------`r`n"

$AbortHomeDrive = $false
$UserFoldersDirectory = "\\mjemelfs2\user$"
$HomeDrivePath = $UserFoldersDirectory + "\" + $SAM
$Domain = "vicmje"

# Create home directory
try
{ 
    New-Item -Path $UserFoldersDirectory -Name $SAM -ItemType Directory -ErrorAction Stop > $null
}
catch
{
    If ($_.CategoryInfo.Category -eq "ResourceExists")
    {
        write-warning "User's home folder already exists"
        if (!(PromptForConfirmation "`r`nAn existing folder has been found at $($HomeDrivePath)`r`n`r`nAre you sure you want to continue with this folder?`r`n(If you choose 'No', you will need to set up the user's home drive manually)"))
        {
            $AbortHomeDrive = $false
        }

        Write-Output ""
    }
    else
    {
        $AbortHomeDrive = $true
        WriteNewestErrorMessage -LogType ERROR -LogString "Could not create user's home directory."
    }    
}

# Get the current permissions of the home drive folder
if (!$AbortHomeDrive)
{
    try
    {
        $ACL = Get-Acl $HomeDrivePath
    }
    catch
    {
        $AbortHomeDrive = $true
        WriteNewestErrorMessage -LogType ERROR -LogString "Could not retrieve the permissions on the new home drive."
    }
}

# The new NTFS permissions rule parameters 
if (!$AbortHomeDrive)
{
    try
    {
        $RuleParameters = @(
            "$($Domain)\$($SAM)"
            "FullControl"
            ,@(
                "ContainerInherit"
                "ObjectInherit"
            )
            "None"
            "Allow"
        )

        # Add the rule to the current permissions list
        $Rule = New-Object `
           -TypeName System.Security.AccessControl.FileSystemAccessRule `
           -ArgumentList $RuleParameters

        $ACL.SetAccessRule($Rule) 
    }
    catch
    {
        $AbortHomeDrive = $true
        WriteNewestErrorMessage -LogType ERROR -LogString "Failed to create home drive permissions."
    }
}

# Set the NTFS permissions on the user's home folder to our new list
if (!$AbortHomeDrive)
{
    try
    {
        Set-Acl -Path $HomeDrivePath -AclObject $ACL
    }
    catch
    {
        $AbortHomeDrive = $true
        WriteNewestErrorMessage -LogType ERROR -LogString "Could not set the NTFS permissions of the user's home folder."
    }
}

# Set home drive on the user's profile
if (!$AbortHomeDrive)
{
    try
    {
        $NewUser | Set-ADUser -HomeDrive "H:" -HomeDirectory $HomeDrivePath
    }
    catch
    {
        $AbortHomeDrive = $true
        WriteNewestErrorMessage -LogType ERROR -LogString "Could set home drive in user's AD account."
    }
}

# Report if errors were detected
if ($AbortHomeDrive)
{
    $ErrorsPresent = $True
    Write-Warning "`r`nErrors were detected during while creating user's home drive. Please correct manually."
    Start-Sleep -Seconds 3
}
else
{
    Write-Output "- Successfully set up user's home drive"
}

<#
----------
Set up User's Mailbox
----------
#>

Write-Output "`r`n----------`r`nEnabling User's Mailbox`r`n----------`r`n"

try 
{
    Enable-Mailbox -identity $SAM -alias $Mail > $null
    Write-Output "- Successfully enabled user mailbox."
}
catch
{
    WriteNewestErrorMessage -LogType ERROR -LogString "Could not enable mailbox."
}

<#
----------
Finishing Tasks
----------
#>

Write-Output "`r`n- User creation has completed.`r`n"


if ($ErrorsPresent)
{
    Write-Warning "`r`n`r`nErrors were encountered - please double check account details"
}