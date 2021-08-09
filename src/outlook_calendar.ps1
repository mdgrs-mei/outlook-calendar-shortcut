# process.ps1 must be included before executing this file.

enum CalendarViewMode
{
    Day = 0
    Week = 1
    Month = 2
    WorkWeek = 4
    Default = 255
}

class OutlookCalendar
{
    $folderPath
    $folderName
    $outlook
    $namespace
    $folder

    [String] Init($folderPath)
    {
        $this.folderPath = $folderPath
        $this.folderName = $folderPath.Substring($folderPath.LastIndexOf("\")+1)
        return $this.InitOutlook()
    }

    [void] Term()
    {
        if (-not $this.IsOutlookValid())
        {
            return
        }
        [system.runtime.interopservices.marshal]::releasecomobject($this.outlook)
    }

    [String] InitOutlook()
    {
        if (-not $this.IsOutlookValid())
        {
            $this.outlook = New-Object -ComObject Outlook.Application
            if (-not $this.IsOutlookValid())
            {
                return "Failed to get Outlook."
            }
        }

        $this.namespace = $this.outlook.GetNamespace("MAPI")
        $this.folder = [OutlookCalendar]::FindFolder($this.namespace.Folders, $this.folderPath)

        if (-not $this.IsFolderValid())
        {
            if (-not $this.folder)
            {
                return "Failed to find folder [{0}]." -f $this.folderPath
            }
            return "Folder is not valid [{0}]." -f $this.folderPath
        }

        return ""
    }

    static [Object] FindFolder($folders, $folderPath)
    {
        foreach ($folder in $folders)
        {
            if ($folder.FolderPath -and ($folder.FolderPath.ToString() -eq $folderPath))
            {
                return $folder
            }

            $f = [OutlookCalendar]::FindFolder($folder.Folders, $folderPath)
            if ($f)
            {
                return $f
            }
        }
        return $null
    }

    [boolean] IsOutlookValid()
    {
        return $this.outlook -and $this.outlook.Name
    }

    [boolean] IsFolderValid()
    {
        return $this.folder.Name
    }

    [String] GetName()
    {
        return $this.folderName
    }

    [String] InitOutlookIfNotValid()
    {
        if ((-not $this.IsOutlookValid()) -or (-not $this.IsFolderValid()))
        {
            return $this.InitOutlook()
        }
        return ""
    }

    [int] GetTodaysRemainingItemCount()
    {
        $errorItemCount = -1
        if (-not $this.IsFolderValid())
        {
            return $errorItemCount
        }

        try 
        {
            $items = $this.GetTodaysRemainingItems()
            if (-not $items)
            {
                return 0;
            }

            # $items.Count should not be used when IncludeRecurrences is True.
            # https://docs.microsoft.com/en-us/office/vba/api/outlook.items.includerecurrences
            $count = 0
            foreach ($item in $items)
            {
                ++$count
            }
            return $count
        }
        catch
        {
            Write-Host "GetTodaysItemCount failed. [$PSItem]"
            return $errorItemCount
        }
    }

    [Object] GetTodaysRemainingItems()
    {
        if (-not $this.IsFolderValid())
        {
            return $null
        }

        $now = Get-Date
        $endOfToday = $now.AddDays(1).Date
        $nowString = $now.ToString("g")
        $endOfTodayString = $endOfToday.ToString("g")
        $query = "[Start] < '$endOfTodayString' And [End] > '$nowString'"

        $items = $this.folder.Items
        $items.IncludeRecurrences = $true
        $items.Sort("[Start]")

        return $items.Restrict($query)
    }

    [void] CreateNewAppointment()
    {
        if (-not $this.IsFolderValid())
        {
            return
        }

        $olAppointmentItem = 1
        $item = $this.folder.Items.Add($olAppointmentItem)
        $item.Display()
        FocusApp "outlook.exe"
    }

    [void] Focus([CalendarViewMode]$viewMode)
    {
        if (-not $this.IsFolderValid())
        {
            return
        }

        $explorer = $this.outlook.ActiveExplorer()
        if (-not $explorer)
        {
            $olFolderInbox = 6
            $inbox = $this.namespace.GetDefaultFolder($olFolderInbox)
            if ($inbox)
            {
                $inbox.Display()
            }
            $explorer = $this.outlook.ActiveExplorer()
        }

        if (-not $explorer)
        {
            return
        }

        try
        {
            $calendarModule = $explorer.NavigationPane.Modules.GetNavigationModule(1)  # 1:olModuleCalendar
            $explorer.NavigationPane.CurrentModule = $calendarModule

            foreach ($group in $calendarModule.NavigationGroups)
            {
                foreach ($navigationFolder in $group.NavigationFolders)
                {
                    if ($navigationFolder.Folder.FolderPath -eq $this.folderPath)
                    {
                        $navigationFolder.IsSelected = $true
                    }
                    else
                    {
                        $navigationFolder.IsSelected = $false
                    }
                }
            }

            $view = $explorer.CurrentView
            # ViewType 2:calendar
            if (($view.ViewType -eq 2) -and ($viewMode -ne [CalendarViewMode]::Default))
            {
                $view.CalendarViewMode = [int]$viewMode
                $view.Save()
            }

            FocusApp "outlook.exe"
        }
        catch 
        {
            Write-Host "Focus failed. [$PSItem]"
        }
    }
}