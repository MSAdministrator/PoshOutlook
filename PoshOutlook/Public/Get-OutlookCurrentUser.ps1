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
function Get-OutlookCurrentUser
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

        try
        {
            $TempObject = $Outlook.Application.Session.CurrentUser.AddressEntry.GetExchangeUser()
        }
        catch
        {
            Write-Error -Message 'Unable to gather information about the curent Outlook user'
        }

        $props = [ordered]@{
                Name = $TempObject.Name
                FirstName = $TempObject.FirstName
                LastName = $TempObject.LastName
                Alias = $TempObject.Alias
                SMTPAddress = $TempObject.PrimarySmtpAddress
                CompanyName = $TempObject.CompanyName
                Department = $TempObject.Department
                OfficeLocation = $TempObject.OfficeLocation
                JobTitle = $TempObject.JobTitle
                AssistantName = $TempObject.AssistantName
                BusinessPhone = $TempObject.BusinessTelephoneNumber
                MobileNumber = $TempObject.MobileTelephoneNumber
                StreetAddress = $TempObject.StreetAddress
                City = $TempObject.City
                State = $TempObject.StateOrProvince
                PostalCode = $TempObject.PostalCode
                Type = $TempObject.Type
                DisplayType = $TempObject.DisplayType
                AddressEntryUserType = $TempObject.AddressEntryUserType
        }

        $ReturnObject = New-Object -TypeName PSCustomObject -Property $props
    }
    End
    {
        Write-Verbose -Message 'Returning Posh.Outlook.Account custom object'

        Add-ObjectDetail -InputObject $ReturnObject -TypeName Posh.Outlook.Account -DefaultProperties Name,Department,JobTitle
    }
}


        