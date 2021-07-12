<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-OutlookAddIns
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param()

    Begin
    {
        $ReturnObject = @()
    }
    Process
    {
        Write-Verbose -Message 'Gathering all installed COM AddIns'

        for ($i = 1; $i -le ($Outlook.Application.COMAddIns).count; $i++)
        {
            Write-Debug -Message 'Generating COM AddIns properties'

            $props = @{
                Name = ($Outlook.Application.COMAddIns[$i]).Description
                Connection = ($Outlook.Application.COMAddIns[$i]).Connect
                GUID = ($Outlook.Application.COMAddIns[$i]).Guid
            }

            Write-Debug -Message 'Creating COM AddIns Object'
            $TempObject = New-Object -TypeName PSCustomObject -Property $props
            $ReturnObject += $TempObject
        }
    }
    End
    {
        Write-Verbose -Message 'Returning Posh.Outlook custom object'

        Add-ObjectDetail -InputObject $ReturnObject -TypeName Posh.Outlook
    }
}