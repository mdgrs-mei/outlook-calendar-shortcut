# process.ps1 must be included before executing this file.

class ActionGenerator
{
    $actionTable = @{}

    [void] Init($calendar)
    {
        $this.actionTable = @{
            "FocusOnCalendar" = {
                $calendar.Focus([CalendarViewMode]::Default)
            }.GetNewClosure()

            "FocusOnToday" = {
                $calendar.Focus([CalendarViewMode]::Day)
            }.GetNewClosure()

            "FocusOnThisWeek" = {
                $calendar.Focus([CalendarViewMode]::Week)
            }.GetNewClosure()

            "FocusOnThisMonth" = {
                $calendar.Focus([CalendarViewMode]::Month)
            }.GetNewClosure()
        }
    }

    [void] Term()
    {
    }

    [Object] CreateAction($name)
    {
        $class = $this

        $block = {
            $class.ExecuteAction($name)
        }.GetNewClosure()

        return $block
    }

    [void] ExecuteAction($name)
    {
        try
        {
            Write-Host ("Action: {0}" -f $name)
            $block = $this.actionTable[$name]
            if ($block)
            {
                $block.Invoke()
            }
        }
        catch
        {
            Write-Host "Action failed. [$PSItem]"
        }
    }
}

