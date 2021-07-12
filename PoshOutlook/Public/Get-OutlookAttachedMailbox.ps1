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
function Get-OutlookAttachedMailbox
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param()

    Begin
    {
        Write-Verbose -Message 'Gathering all attached mailboxes'

        $ReturnObject = @()
    }
    Process
    {
        Write-Debug -Message 'Looping through all Outlook mailboxes'

        foreach ($item in $Outlook.Session.Stores)
        {
            Write-Verbose -Message 'Creating properties for custom object'

            $props = @{
                Name = $item.DisplayName
                CachedMode = $item.IsCachedExchange
                SearchEnabled = $item.IsInstantSearchEnabled
                ConversationView = $item.IsConversationEnabled
            }

            $TempObject = New-Object -TypeName PSCustomObject -Property $props
            $ReturnObject += $TempObject
        }
    }
    End
    {
        Write-Verbose -Message 'Returning Posh.Outlook custom object'

        Add-ObjectDetail -InputObject $ReturnObject -TypeName Posh.Outlook.Mailbox
    }
}