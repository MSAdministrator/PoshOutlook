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
function Get-OutlookSessionInfo
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param()

    Begin
    {
        $ReturnObject = @()

        Write-Verbose -Message 'Creating custom object properties in Get-OutlookSessionInfo'

        Write-Debug -Message 'Creating custom object properties'

        $props = @{
            ProfileName = $Outlook.Session.CurrentProfileName
            FilePath = $Outlook.Session.DefaultStore.FilePath
            CachedMode = $Outlook.Session.DefaultStore.IsCachedExchange
            ConversationView = $Outlook.Session.DefaultStore.IsConversationEnabled
            OfflineStatus = $Outlook.Application.Session.Offline
            ExchangeConnectionMode = $Outlook.Application.Session.ExchangeConnectionMode
            Namespace = ($Outlook.GetNameSpace("MAPI")).Stores
            ProductCode = $Outlook.Application.ProductCode
        }
    }
    Process
    {
        Write-Debug -Message 'Creating custom object'

        $ReturnObject = New-Object -TypeName PSCustomObject -Property $props
    }
    End
    {
        Write-Verbose -Message 'Returning custom object'

        Add-ObjectDetail -InputObject $ReturnObject -TypeName Posh.Outlook.Sesssion
    }
}