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
function Get-OutlookVersions
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Type of Outlook
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        [ValidateSet('Outlook','Exchange')]
        $Type
    )

    Begin
    {
        $Version = @()

        Write-Debug -Message 'Attempting to verify which version is being called'

        try
        {
            if ($Type -eq 'Exchange')
            {
                $Version = [Version]$Outlook.Application.Session.ExchangeMailboxServerVersion
            }
            elseif ($Type -eq 'Outlook')
            {
                $Version = [version]$Outlook.Version
            }
        }
        catch
        {
            Write-Error -Message "Unable to get version information for $Type"
        }

        Write-Verbose -Message 'Creating properties for custom object'

        $props = [ordered]@{
            Major = $Version.Major
            Minor = $Version.Minor
            Build = $Version.Build
            Revision = $Version.Revision
            MajorRevision = $Version.MajorRevision
            MinorRevision = $Version.MinorRevision
        }
    }
    Process
    {
        Write-Verbose -Message 'Creating Custom Object'

        try
        {
            $TempObject = New-Object -TypeName PSCustomObject -Property $props
        }
        catch
        {
            Write-Error -Message 'Unable to create PSCustomObject'
        }
    }
    End
    {
        Add-ObjectDetail -InputObject $TempObject -TypeName Posh.Outlook.Version
    }
}