function Get-MirrorPOBox 
{
    [CmdletBinding()]
    [OutputType([string])]    
    param (
        # The account to retrieve the PO Boxvalue  from
        [Parameter(Mandatory=$true)]
        [string]
        $Identity
    )

    try 
    {
        $POBox = (get-aduser -Identity $Identity -Properties "extensionAttribute1").extensionattribute1
    } 
    catch
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Could not retrieve the mirror user for the PO Box"
    }

    return $POBox
}

function Set-POBox
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Identity, 

        # Parameter help description
        [Parameter()]
        [string]
        $POBox
    )

    if (!$POBox)
    {
        Write-Warning "POBox value is blank. Continuing"
        return
    }

    try 
    {
        Set-ADUser -Identity $Identity -Add @{extensionAttribute1 = $POBox}
        Write-Host "- Set POBox (extensionAttribute1)"
    }
    catch 
    {
        Write-NewestErrorMessage -LogType ERROR -CaughtError $_ -LogToFile $true -LogString "Something went wrong while setting the PO Box (extensionAttribute1)"
    }
}