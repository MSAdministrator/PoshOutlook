Function Get-OutlookCalendarItem 
{  
  
  
<#  
    .SYNOPSIS  
        A function to retreive Outlook Calender items between two dates.   
        Returns PSobjects containing each item.  
    .DESCRIPTION  
        The function opens one's outlook calender, then retrieves the items.   
        The function takes 2 parameter: start, end - items are returned that   
        start betweween these two dates.  
    .NOTES  
        File Name  : Get-OLCalendarItem  
        Author     : Thomas Lee - tfl@psp.co.uk  
        Requires   : PowerShell Version 3.0
        Modified By: Josh Rickard
        Modified Date: 11/24/2016
    .LINK  
        This script posted to:  
            http://pshscripts.blogspot.com/2013/10/get-olcalendaritem.html 
    .Link
        http://blogs.technet.com/b/heyscriptingguy/archive/2011/05/24/use-powershell-to-export-outlook-calendar-information.aspx
      
    .EXAMPLE  
        Left as an exercise for the reader       
 
#>  
  
    [CmdletBinding()]  
    Param (
        [datetime]$start = '1/1/1900',  

        [datetime]$end = $(Get-Date -UFormat %m/%d/%Y)
    )  
  
    Begin
    {
        Write-Verbose "Returning Outlook Calendar Items"  
        Write-Verbose "Dates Start = $Start;  End = $End"

        Try
        {
            $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type] 
        }
        Catch
        {
            $_
        }

        $ReturnObject = @()

        $Calendar = $namespace.getDefaultFolder($olFolders::olFolderCalendar) 

        Write-Verbose "There are $($calendar.items.count) items in the calender in total"  
    }
    Process
    {
        # Now return psobjects for all items between 2 dates  
        ForEach ($citem in ($Calendar.Items | sort start)) {  
            Write-Verbose "Processing ($Citem | Out-String)"  
  
            If ($citem.start -ge $start -and $citem.start -LE $end) {   
            
                #$CItem | gm

                $CalHT =[ordered]  @{  
                    Subject          =  $($Citem.Subject)  
                    Location         =  $($Citem.Location);  
                    Start            =  $(Get-Date $Citem.Start);  
                    StartUTC         =  $(Get-Date $Citem.StartUTC);                                    
                    End              =  $(Get-Date $Citem.End);  
                    EndUTC           =  $(Get-Date $Citem.EndUTC);  
                    Duration         =  $($Citem.Duration);  
                    Importance       =  $($Citem.Importance);  
                    IsRecurring      =  $($Citem.IsRecurring);  
                    AllDayEvent      =  $($citem.AllDayEvent);  
                    Sensitivity      =  $($Citem.Sensitivity);  
                    ReminderSet      =  $($Citem.ReminderSet);  
                    CreationTime     =  $($Citem.CreationTime);  
                    LastModificationTime = $($Citem.LastModificationTime);  
                    Categories       =  $($CItem.Categories);
                    Body             =  $($Citem.Body);  
                }  
  
                # Write the item out as a custom item  
                $TempObject = New-Object PSobject -Property $CalHT  
                $ReturnObject += $TempObject  
  
            }  
        } #End of foreach  
    }
    End
    {
        Add-ObjectDetail -InputObject $ReturnObject -TypeName Posh.Outlook.Calendar
    }
}  # End of function    
