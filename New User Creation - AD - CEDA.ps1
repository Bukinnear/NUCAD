Write-Output "Loading - please wait`r`n"

Import-Module ActiveDirectory

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
    # This is part of a text prompt for the mirrored user's account
    [string]$Script:UsernameFormat = "Firstname Lastname = LastnameF"
    
    # Email 
    # (Only used below in primary/alias SMTP addresses, this can change later in the script)
    # LDAP Fields:   
    # - N/A 
    $Script:Mail = $FirstName+"."+$LastName

    # User logon name/User Principle Name 
    # WARNING: This is separate from $Mail, and should not be set the same!
    # LDAP Fields:
    # - UserPrincipleName (User Logon Name)
    $Script:UPN = "$($FirstName+"."+$LastName)@CEDA.com.au"
    
    # Pre-windows 2000 Logon name
    # LDAP Fields:
    # - SamAccountName    
    $Script:SAM = $LastName+$FirstName[0]
    
    # Remove any hyphens from the username
    $SAM = $SAM -replace "-", ""
    $UPN = $UPN -replace "-", ""

    GenerateProxyAddresses
}

# Prepares email address variables
function GenerateProxyAddresses()
{            
    # Used as the primary SMTP address (this is set with the proxy addresses below)
    # - Mail (EmailAddress)
    # - PrimarySmtpAddress (not explicitly set, but by extension of proxy addresses
    $Script:EmailAddress = "$($Mail)@CEDA.com.au"

    # Array of mail aliases
    $Script:ProxyAddresses = "SMTP:$($EmailAddress)", "smtp:$($Mail)@CEDA.mail.onmicrosoft.com"
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
            $msgConfimation = Read-Host "`r`n$($Message) Y/N"
        }

        $msgConfimation = $msgConfimation.ToLower()    
    
        switch ($msgConfimation)
        {
            {$_ -in 'yes','y'}{return $true}
            {$_ -in 'no','n'}{return $false}
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

<#
----------
# Get new user details
----------
#>

do 
{
    $ReadyToCreate = $false

    # Get the user's name
    do
    {
        Write-Warning "It is not recommended to create service accounts with this tool.`r`n"

        [string]$FirstName = Read-Host "Please enter a First Name"
        [string]$LastName = Read-Host "Please enter a Last Name"

        #Remove whitespaces
        $Firstname = $Firstname.Trim()
        $Lastname = $Lastname.Trim()
        
        # Spacing
        Write-Output ""

        If(PromptForConfirmation "Would you like to fix/standardise the capitalisation of this name?`r`n(Choose yes if you are not sure)")
        {
            # Capitalise (only) the first letter of the names
            $FirstName = (Get-Culture).TextInfo.ToTitleCase($FirstName.ToLower())
            $LastName = (Get-Culture).TextInfo.ToTitleCase($LastName.ToLower())
        }
    } while (!(PromptForConfirmation "Is this name correct?`r`n`r`n$($FirstName) $($LastName)"))
    
    # Get the Job Title
    do 
    {
        $JobTitle = Read-Host "Please enter a Job Description"
    } while (!(PromptForConfirmation "Is this correct?`r`n`r`n$($JobTitle)"))

    # Get the phone number
    do 
    {
        $PhoneNumber = Read-Host "Please enter a phone number (Leave blank for none)"
    } while (!(PromptForConfirmation "Is this correct?`r`n`r`n$($PhoneNumber)"))

    # Spacing
    Write-Output ""

    # Set the default password
    $Password = "`$Password99!"

    # Allow for a custom password
    if (!(PromptForConfirmation "Would you like to set the default password?`r`n(Choose 'No' to create your own)`r`n`r`n$($Password)"))
    {
        do
        {
            $Password = Read-Host "Please enter a Password"
        } while ($Password -eq "" -or $Password -eq $null)
    }

    # Set client specific parameters, such as email, and logon name
    SetClientSpecificParams -FirstName $FirstName -LastName $LastName
        
    while (!(PromptForConfirmation "The primary SMTP/email address as been set to:`r`n`r`n$($EmailAddress)`r`n`r`nIs this correct?"))
    {        
        $Mail = Read-Host "Enter the email address (Only enter the text BEFORE the '@' sign)"
        GenerateProxyAddresses  
    } 

    # Spacing
    Write-Output ""

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

do
{
    $Script:PasswordInvalid = $False

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
        
        WriteNewestErrorMessage -LogType ERROR -LogString "Failed to create new user. Password is invalid."

        try 
        {
            # Try to get the new account - this will stop here if the account cannot be found
            Get-ADUser -Identity $SAM > $null

            Write-Warning "`r`nRemoving the account and starting again`r`n"

            # Remove the account if it has been created.
            Remove-ADUser -Identity $SAM -Confirm:$false
            
        }
        catch {}

        $Password = Read-Host "Please enter a password"
    }
    catch
    {    
        WriteNewestErrorMessage -LogType ERROR -LogString "Failed to create new user. Exiting"
        return
    }
} while ($Script:PasswordInvalid)

Start-Sleep -Seconds 3

<#
----------
Populate the new account's properties
----------
#>

# Get the new user account
try 
{
    $NewUser = Get-ADUser -Identity $SAM -Properties *
}
catch
{
    WriteNewestErrorMessage -LogType ERROR -LogString "Failed to create/get new account. Exiting"
    return
}

Write-Output "`r`nSuccessfully created the new user account. Populating account details."

# Copy mirror user's details
try
{
    $NewUser | Set-ADUser -Manager $MirrorUser.manager -State $MirrorUser.st -Country $MirrorUser.c -PostalCode $MirrorUser.postalCode -StreetAddress $MirrorUser.streetAddress -City $MirrorUser.l -Office $MirrorUser.physicalDeliveryOfficeName -HomePage $MirrorUser.HomePage
    Write-Output "`r`n- Address and manager have been set."
}
catch
{
    WriteNewestErrorMessage -LogType ERROR -LogString "Could not set some parameters - please double check the account's address and manager"
    $ErrorsPresent = $True
}

# Set account's email and proxy addresses
try
{
    $NewUser | Set-ADUser -Add @{ProxyAddresses = $ProxyAddresses}
    $NewUser | Set-ADUser -add @{MailNickName = $SAM}
    $NewUser | Set-ADUser -add @{Mail = $EmailAddress}
    Write-Output "- Email/Proxy addresses have been set."
}
catch
{
    WriteNewestErrorMessage -LogType ERROR -LogString "Could not set some parameters - please double check the account's email and proxy addresses"
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

Write-Output "`r`nUser creation has completed.`r`n"

if (!(PromptForConfirmation "Would you like to run a sync to O365?"))
{
    return
}

<#
----------
Sync to O365
----------
#>

if (!$ErrorsPresent)
{
    Write-Output "`r`n----------`r`nBeginning sync to O365`r`n----------`r`n"
    Start-sleep -Seconds 3
    Start-ADSyncSyncCycle -PolicyType Delta
}
else
{
    Write-Warning "Errors were encountered, Sync to O365 has not been run - please check all account details and run sync manually with:`r`n`r`nStart-ADSyncSyncCycle -PolicyType Delta"
}