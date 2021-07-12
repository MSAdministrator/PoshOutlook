#requires -Version 2

if (Get-Variable -Name Outlook -ErrorAction SilentlyContinue){
    Remove-Variable -Name Outlook -Force -ErrorAction SilentlyContinue
}

[Microsoft.Office.Interop.Outlook.Application] $Outlook = New-Object -ComObject Outlook.Application

if (Get-Variable -Name Namespace -ErrorAction SilentlyContinue){
    Remove-Variable -Name Namespace -Scope Global -Force -ErrorAction SilentlyContinue
}

try
{
    New-Variable -Name Namespace -Value $(($Outlook.GetNameSpace("MAPI")).Stores) -Description 'A global variable for the Outlook namespace' -Scope Global
}
catch
{
    Write-Error -Message 'Unable to create the global Namespace variable'
}


#Get public and private function definition files.
$Public  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -Recurse -ErrorAction SilentlyContinue )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -Recurse -ErrorAction SilentlyContinue )

#Dot source the files
Foreach($import in @($Public + $Private))
{
    Try
    {
        . $import.fullname
    }
    Catch
    {
        Write-Error -Message "Failed to import function $($import.fullname): $_"
    }
}


Export-ModuleMember -Variable Outlook,Namespace
Export-ModuleMember -Function $Public.Basename
