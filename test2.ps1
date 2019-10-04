$ErrorActionPreference = "Stop"

# Path to new user creation module
$Script:ScriptPath = split-path -parent $MyInvocation.MyCommand.Definition

# For debugging purposes
if (($ScriptPath -eq "") -or ($null -eq $ScriptPath))
{
    $ScriptPath = "C:\Code\PS\Testing\New Users"
}

# Import the new user creation module
$test = Import-Module -Name "$($ScriptPath)\NUC-AD" -Force -PassThru

$Script:UserFolderDirectory = "C:\temp\"
$Script:Domain = "kiandra"

$HomeDrive = New-UserFolder `
    -SamAccountName  "Jared.kinnear" `
    -Domain $Domain `
    -ParentFolderPath $UserFolderDirectory `
    -FolderName "Test3"