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
function Search-OutlookAttachment
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$False,
                   ValueFromPipeline=$true,
                   Position=1)]
        [System.string]$Filename,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   Position=1)]
        [string]$Folder = 'InBox'


    )

    Begin
    {
        $ReturnObject = @()
    }
    Process
    {
        Write-Verbose -Message 'Gathering information about the current Outlook user'

        try
        {
            $Email = $NameSpace.Folders.Item(1).Folders.Item($Folder).Items

            foreach ($e in $Email)
            {
                $e.Attachments | foreach { 
                    if ($item){
                        if (($_.FileName).Contains($Item))
                        { 
                            $props = [ordered]@{
                                Name = $_.DisplayName
                                FileName = $_.FileName
                                Size = $_.Size
                            }

                        $ReturnObject = New-Object -TypeName PSCustomObject -Property $props
                        Add-ObjectDetail -InputObject $ReturnObject -TypeName Posh.Outlook.Attachment
                        }
                    }
                    else{
                        $props = [ordered]@{
                            Name = $_.DisplayName
                            FileName = $_.FileName
                            Size = $_.Size
                        }

                        $ReturnObject = New-Object -TypeName PSCustomObject -Property $props
                        Add-ObjectDetail -InputObject $ReturnObject -TypeName Posh.Outlook.Attachment
                    }
                }
            }
        }
        catch
        {
            Write-Error -Message 'Unable to search Outlook for that attachment'
        }
    }
    End
    {
        # Intentionally left blank
    }
}