# process.ps1 must be included before executing this file.

class ActionGenerator
{
    $actionTable = @{}

    [void] Init($calendar)
    {
        $this.actionTable = @{
            "FocusOnCalendar" = {
                $calendar.Focus([CalendarViewMode]::Default, $null)
            }.GetNewClosure()

            "FocusOnToday" = {
                $calendar.Focus([CalendarViewMode]::Day, $null)
            }.GetNewClosure()

            "FocusOnThisWeek" = {
                $calendar.Focus([CalendarViewMode]::Week, $null)
            }.GetNewClosure()

            "FocusOnThisWorkWeek" = {
                $calendar.Focus([CalendarViewMode]::WorkWeek, $null)
            }.GetNewClosure()

            "FocusOnNextNDays" = {
                param($days)
                $calendar.Focus([CalendarViewMode]::MultiDay, $days)
            }.GetNewClosure()

            "FocusOnThisMonth" = {
                $calendar.Focus([CalendarViewMode]::Month, $null)
            }.GetNewClosure()

            "CreateNewAppointment" = {
                $calendar.CreateNewAppointment()
            }.GetNewClosure()
        }
    }

    [void] Term()
    {
    }

    [Object] CreateAction($action)
    {
        $class = $this

        $block = {
            $class.ExecuteAction($action)
        }.GetNewClosure()

        return $block
    }

    [void] ExecuteAction($action)
    {
        try
        {
            $actionName = $action[0]
            $actionArgs = $action[1..($action.Count-1)]

            Write-Host ("Action: {0}" -f $actionName)
            $block = $this.actionTable[$actionName]
            if ($block)
            {
                Invoke-Command $block -ArgumentList $actionArgs
            }
        }
        catch
        {
            Write-Host "Action failed. [$PSItem]"
        }
    }
}

