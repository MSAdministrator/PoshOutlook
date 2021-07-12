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
function Get-OutlookGAL
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
        Write-Verbose -Message 'Gathering the Global Address List'

        try
        {
            $Entries = $Outlook.Session.GetGlobalAddressList().AddressEntries
        }
        catch
        {
            Write-Error -Message 'Unable to gather global address list from Outlook'
        }

        Write-Debug -Message 'Iterating through list of global address entries'

        foreach ($entry in $Entries)
        {
            Write-Verbose -Message "Processing $Entry.Name"

            $User = $entry.GetExchangeUser()

            $props = [ordered]@{
                Name = $User.Name
                FirstName = $User.FirstName
                LastName = $User.LastName
                Alias = $User.Alias
                SMTPAddress = $User.PrimarySmtpAddress
                CompanyName = $User.CompanyName
                Department = $User.Department
                OfficeLocation = $User.OfficeLocation
                JobTitle = $User.JobTitle
                AssistantName = $User.AssistantName
                BusinessPhone = $User.BusinessTelephoneNumber
                MobileNumber = $User.MobileTelephoneNumber
                StreetAddress = $User.StreetAddress
                City = $User.City
                State = $User.StateOrProvince
                PostalCode = $User.PostalCode
                Type = $User.Type
                DisplayType = $User.DisplayType
                AddressEntryUserType = $User.AddressEntryUserType
        }

            $TempObject = New-Object -TypeName PSCustomObject -Property $props
            $ReturnObject += $TempObject
        }
    }
    End
    {
        Write-Verbose -Message 'Returning Posh.Outlook.GAL custom object'

        Add-ObjectDetail -InputObject $ReturnObject -TypeName Posh.Outlook.GAL -DefaultProperties Name,Department,JobTitle
    }
}