Write-Output "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

# Path to new user creation module
$Script:ScriptPath = (split-path -parent $MyInvocation.MyCommand.Definition)
Import-Module -Name "$($ScriptPath)\NUC-AD" -Force
#Import-Module -Name "PATH\NUC-AD" -Force