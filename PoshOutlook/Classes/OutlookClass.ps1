[Reflection.Assembly]::LoadWithPartialname(“Microsoft.Office.Interop.Outlook”) | out-null

class Outlook
{
    $OutlookVersion
    $ExchangeVersion
    hidden $NameSpace
    [string]$ProfileName
    [string]$FilePath
    [bool]$CachedMode
    [bool]$InstantSearchEnabled
    [bool]$ConversationView
    [bool]$OfflineStatus
    [int]$ExchangeConnectionMode
    $Accounts
    $AttachedMailbox
    $Instance
    $ComAddins
    $ProductCode
    $GlobalAddressList
    hidden $RootFolderNames
    [System.Object]$OutlookFolders = @()
    [System.Object]$CurrentUser = @()

    Outlook (){ }

    Outlook ([Microsoft.Office.Interop.Outlook.ApplicationClass]$Outlook)
    {
        $this.ProfileName = $Outlook.Session.CurrentProfileName
        $this.FilePath = $Outlook.Session.DefaultStore.FilePath
        $this.CachedMode = $Outlook.Session.DefaultStore.IsCachedExchange
        $this.ConversationView = $Outlook.Session.DefaultStore.IsConversationEnabled
        $this.OfflineStatus = $Outlook.Application.Session.Offline

        $this.ExchangeConnectionMode = $Outlook.Application.Session.ExchangeConnectionMode
        
        $this.ExchangeVersion = $this.GetVersion($Outlook, 'Exchange')
        $this.OutlookVersion = $this.GetVersion($Outlook, 'Outlook')

        $this.Instance = $Outlook
        $this.Namespace = ($Outlook.GetNameSpace("MAPI")).Stores
       
        $this.ProductCode = $Outlook.Application.ProductCode

        $this.GetAccounts($Outlook)
        $this.GetAttachedMailbox($Outlook)
        $this.GetComAddins($Outlook)
        $this.GetCurrentUser($Outlook)
        $this.GetFolders($Outlook)
        $this.GetGlobalAddressList($Outlook)
    }

    hidden [System.Object] GetVersion ($Outlook, $Type)
    {
        $Version = @()

        if ($type -eq 'Exchange')
        {
            $Version = [Version]$Outlook.Application.Session.ExchangeMailboxServerVersion
        }
        elseif ($type -eq 'Outlook')
        {
            $Version = [version]$Outlook.Version
        }

        $props = [ordered]@{
            Major = $Version.Major
            Minor = $Version.Minor
            Build = $Version.Build
            Revision = $Version.Revision
            MajorRevision = $Version.MajorRevision
            MinorRevision = $Version.MinorRevision
        }

        $TempObject = New-Object -TypeName PSCustomObject -Property $props
        return $TempObject
    }

    [System.Object] GetAccounts ($Outlook){
        $AccountObject = @()

        foreach ($item in $this.Namespace)
        {
            $props = @{ 
                DisplayName = $item.DisplayName
                UserName = $item.UserName
                SmtpAddress = $item.SmtpAddress
                ProfileName = ($item.Session).CurrentProfileName
            }

            $TempAccountObject = New-Object -TypeName PSCustomObject -Property $props

            $AccountObject += $TempAccountObject
        }

        $this.Accounts = $AccountObject

        return $this.Accounts
    }

    [System.Object] GetAttachedMailbox ($Outlook)
    {
        $ReturnObject = @()

        foreach ($item in $Outlook.Session.Stores)
        {
            $props = @{
                Name = $item.DisplayName
                CachedMode = $item.IsCachedExchange
                SearchEnabled = $item.IsInstantSearchEnabled
                ConversationView = $item.IsConversationEnabled
            }

            $TempObject = New-Object -TypeName PSCustomObject -Property $props
            $ReturnObject += $TempObject
        }

        $this.AttachedMailbox = $ReturnObject
        return $this.AttachedMailbox
    }

    [System.Object] GetComAddins ($Outlook)
    {
        $ReturnObject = @()

        for ($i = 1; $i -le ($Outlook.Application.COMAddIns).count; $i++)
        {
       
            $props = @{
                Name = ($Outlook.Application.COMAddIns[$i]).Description
                Connection = ($Outlook.Application.COMAddIns[$i]).Connect
                GUID = ($Outlook.Application.COMAddIns[$i]).Guid
            }

            $TempObject = New-Object -TypeName PSCustomObject -Property $props
            $ReturnObject += $TempObject
        }

        $this.ComAddins = $ReturnObject
        return $this.ComAddins
    }

    [System.Object] GetGlobalAddressList ($Outlook)
    {
        $ReturnObject = @()

        $Entries = $Outlook.Session.GetGlobalAddressList().AddressEntries

        foreach ($entry in $Entries)
        {
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

        $this.GlobalAddressList = $ReturnObject
        return $this.GlobalAddressList
    }

    [System.Object] GetCurrentUser ($Outlook)
    {
        $TempObject = $Outlook.Application.Session.CurrentUser.AddressEntry.GetExchangeUser()

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

        $this.CurrentUser = New-Object -TypeName PSCustomObject -Property $props
        return $this.CurrentUser

    }

    [system.object] GetFolders($Outlook)
    {
        $ReturnObject = @()

        function Get-MailboxFolder($folder)
        {
            $props += @{
                Name = $folder.Name
                Count = $folder.items.count
            }

            $TempObject = New-Object -TypeName PSCustomObject -Property $props
            $this.OutlookFolders += $TempObject

            foreach ($f in $folder.folders)
            {
                Get-MailboxFolder $f
            }
        }

        $mailbox = $this.NameSpace | where {$_.ExchangeStoreType -eq 0}
        $mailbox | Out-String

        $TempRootFolders = $mailbox.GetRootFolder().Folders | Select -Property Name
        $this.RootFolderNames = $TempRootFolders
        $mailbox.GetRootFolder().folders | foreach { Get-MailboxFolder $_ }
        return $this.OutlookFolders
    }
}