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
function Get-OutlookRootFolders
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
        Write-Verbose -Message 'Gathering information about the current Outlook user'

        foreach ($folder in $outlook.Session.Folders){            
              
            foreach($mailfolder in $folder.Folders ) {            
               
                if ($deleted) {if ($($mailfolder.Name) -notlike "Deleted*"){continue} }            
                if ($junk)  {if ($($mailfolder.Name) -notlike "Junk*"){continue} }            
           
   
                $TempObject = New-Object -TypeName PSObject -Property @{            
                    Mailbox = $($folder.Name)            
                    Folder = $($mailfolder.Name)            
                    ItemCount = $($mailfolder.Items.Count)            
                }
                
                $ReturnObject += $TempObject
                 # | select Mailbox, Folder, ItemCount            
            }             
        }            
    }
    End
    {
        Write-Verbose -Message 'Returning Posh.Outlook.Account custom object'

        Add-ObjectDetail -InputObject $ReturnObject -TypeName Posh.Outlook.Folders -DefaultProperties Mailbox,Folder,ItemCount
    }
}