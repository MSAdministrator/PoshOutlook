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
function Get-OutlookAccounts
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param()

    Begin
    {
        Write-Verbose -Message 'Gathering Outlook accounts'

        $AccountObject = @()
    }
    Process
    {
        Write-Debug -Message 'Looping through all Outlook accounts'

        foreach ($item in $(($namespace.Session).Accounts))
        {
            Write-Verbose -Message 'Creating properties for custom object'

            $props = @{ 
                DisplayName = $item.DisplayName
                UserName = $item.UserName
                SmtpAddress = $item.SmtpAddress
                ProfileName = ($item.Session).CurrentProfileName
            }

            Write-Verbose -Message 'Creating PSCustomObject'

            $TempAccountObject = New-Object -TypeName PSCustomObject -Property $props

            $AccountObject += $TempAccountObject
        }
    }
    End
    {
        Write-Verbose -Message 'Returning Posh.Outlook custom object'

        Add-ObjectDetail -InputObject $AccountObject -TypeName Posh.Outlook.Account
    }
}