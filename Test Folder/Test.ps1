Write-Output "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

# Path to new user creation module
$Script:ScriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$Script:ScriptPath = split-path -parent $ScriptPath

# For debugging purposes
if (($ScriptPath -eq "") -or ($null -eq $ScriptPath))
{
    $ScriptPath = "C:\Code\PS\Testing\New Users"
}

# Import the new user creation module
Import-Module -Name "$($ScriptPath)\NUC-AD" -Force
if (!(Initialize-Module)) { return }

