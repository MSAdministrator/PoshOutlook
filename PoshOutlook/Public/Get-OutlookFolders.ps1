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
function Get-OutlookFolders
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param()

    Begin
    {
        $TempObject = @()
    }
    Process
    {
        Write-Verbose -Message 'Gathering information about the current Outlook user'

        function Get-MailboxFolder($folder, $rootFolder)
        {
            $props += [ordered]@{
                RootFolder = $rootFolder.name
                SubFolder = $folder.Name
                Count = $folder.items.count
            }

            $TempObject = New-Object -TypeName PSCustomObject -Property $props
            $TempObject
    
            foreach ($f in $folder.folders)
            {
                Get-MailboxFolder $f $rootFolder
            }
        }

        $nSpace = $Outlook.GetNamespace("MAPI")
        $mailbox = $nSpace.stores | where {$_.ExchangeStoreType -eq 0}
        $mailbox
        $RootFolderNames = $mailbox.GetRootFolder().Folders | Select -Property Name

        $mailbox.GetRootFolder().folders | foreach { 
            foreach ($r in $RootFolderNames){
                if ($r -match $_.Name){ 
                    Get-MailboxFolder $_ $r 
                }
            }
        }   
    }
    End{}
}
