Add-Type -AssemblyName PresentationFramework

$scriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location $scriptDir

. .\settings_helper.ps1
. .\process.ps1
. .\outlook_calendar.ps1
. .\delegate_command.ps1
. .\window.ps1
. .\action_generator.ps1

$settingsPath = $args[0]
. $settingsPath
SetSettingsDirectory $settings $settingsPath

$calendar = [OutlookCalendar]::new()
$calendar.Init($settings.outlook.folderPath)

$windowTitle = "Outlook Calendar"
$window = [Window]::new()
$window.Init(".\window.xaml", $windowTitle, $settings)

$actionGenerator = [ActionGenerator]::new()
$actionGenerator.Init($calendar)

$clickAction = $actionGenerator.CreateAction($settings.clickAction)
$window.SetOnClickedFunction($clickAction)

foreach ($thumbButtonSetting in $settings.thumbButtons)
{
    $button = $window.AddThumbButton($thumbButtonSetting)
    $action = $actionGenerator.CreateAction($thumbButtonSetting.clickAction)
    $button.Command = New-Object DelegateCommand($action)
}

function TimerFunction()
{
    $calendar.InitOutlookIfNotValid()

    if ($settings.todaysRemainingItemsSummary.enable)
    {    
        $eventsSummary = $calendar.GetTodaysRemainingItemsSummary(
            $settings.todaysRemainingItemsSummary.maxItemCount,
            $settings.todaysRemainingItemsSummary.maxItemCharacterCount)
        $window.SetTaskbarItemInfoDescription($eventsSummary)
    }

    if ($settings.overlayIcon.enable)
    {
        $count = $calendar.GetTodaysRemainingItemCount()
        $window.UpdateOverlayCount($count)
    }
}

$window.StartTimerFunction({TimerFunction}, $settings.updateIntervalInSeconds)
$window.ShowDialog()
$window.Term()

$actionGenerator.Term()
$calendar.Term()
