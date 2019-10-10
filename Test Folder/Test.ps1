
function Search-UserAccounts1
{
    Write-Host "`r`n----------"

    for (;;)
    {        
        $Search = Read-Host "Enter a name, or part of a name to search for. leave it blank to continue`r`n"

        if ("" -eq $Search) { return }
        
        $Results = Get-ADUser -Filter "name -like '*$($Search)*'" -Properties SamAccountName, Name, DisplayName, EmailAddress | select SamAccountName, Name, DisplayName, EmailAddress | Format-Table
        Wait-Debugger
        Write-Host "`r`n----------"
        foreach ($Item in $Results) { Write-Host $_ }
        Write-Host "----------`r`n"
    }
}

Search-UserAccounts1

return
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

Search-UserAccounts