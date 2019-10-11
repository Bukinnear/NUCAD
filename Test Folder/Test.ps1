$Test = @{
    K1 = "Value1"
    K2 = "Value2"
    K3 = "Value3"
}

Wait-Debugger
return

Write-Host "Loading - please wait`r`n"
$ErrorActionPreference = "Stop"

# Get the script's file path
$Script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Add this path to the modules search directory
if ($env:PSModulePath -notlike "*$($Script:ScriptPath)*")
{
    $env:PSModulePath += (";" + $ScriptPath)
}

Import-Module -Name "NUC-AD" -Force
if (!(Initialize-Module)) { Write-Error "Could not import NUC-AD module" }

