function TestF 
{
    param (
        [Parameter(
            Mandatory=$true,
            Position=0,
            ValueFromPipelineByPropertyName=$true
        )]
        [string]
        $Para1,

        [Parameter(
            Mandatory=$true,
            Position=1,
            ValueFromPipelineByPropertyName=$true
        )]
        [string]
        $Para2,

        [Parameter(
            Mandatory=$true,
            Position=2,
            ValueFromPipelineByPropertyName=$true
        )]
        [string]
        $Para3
    )

    Write-Host "`r`nPara1: $($Para1)`r`nPara2: $($Para2)`r`nPara3: $($Para3)"
}

$Hash = @{Para1 = "This"; Para2 = "Is"; Para3 = "A test"}

$Hash | TestF

$Hash.Para1