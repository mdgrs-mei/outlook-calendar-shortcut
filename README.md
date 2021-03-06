<div align="center">

# Outlook Calendar Shortcut

[![GitHub license](https://img.shields.io/github/license/mdgrs-mei/outlook-calendar-shortcut)](https://github.com/mdgrs-mei/outlook-calendar-shortcut/blob/main/LICENSE)
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/mdgrs-mei/outlook-calendar-shortcut?label=latest%20release)](https://github.com/mdgrs-mei/outlook-calendar-shortcut/releases/latest)
[![GitHub all releases](https://img.shields.io/github/downloads/mdgrs-mei/outlook-calendar-shortcut/total)](https://github.com/mdgrs-mei/outlook-calendar-shortcut/releases/latest)

Outlook Calendar Shortcut is a Windows taskbar application that works as a shortcut to Outlook's calendar view. It notifies you of today's remaining event count by an overlay badge. Clicking the taskbar icon leads you to the calendar view in Outlook.

![taskbar](./docs/taskbar.gif)

</div>

## Features
- Notifies today's remaining event count
- Customizable icons and badge colors
- Quick access to the Day, Week and Month views with Thumb buttons

## Requirements
- Tested on Windows 10/11 and Powershell 5.1
- Outlook desktop app needs to be installed

## Installation
1. Download and extract the [zip](https://github.com/mdgrs-mei/outlook-calendar-shortcut/releases/latest/download/outlook-calendar-shortcut.zip) or clone this repository anywhere you like
1. Copy and edit `settings.ps1` (See [Settings](#Settings))
1. Run [`tools/create_shortcut.bat`](#toolscreate_shortcutbat) and save the shortcut
1. Run the shortcut

# Settings
You can customize the behavior by a settings file. A sample settings file is placed at [sample/settings.ps1](./sample/settings.ps1).

## Outlook settings

```powershell
outlook = @{
    folderPath = "\\your-email-address@sample.com\calendar-folder-name"
}
```
`folderPath` is a path of the outlook calendar folder that the app monitors. You can list all of your calendar folder paths by running [`tools/list_outlook_calendar_folders.bat`](#toolslist_outlook_calendar_foldersbat).

## Icon image

```powershell
iconPath = ".\icon.png"
```
An icon file used for the title bar. The image is converted to `.ico` file during the shortcut creation and also used as a shortcut icon. `.bmp`, `.png`, `.tif`, `.gif` and `.jpg` with single resolution are supported.

## Overlay icon

![overlay_icon](./docs/overlay_icon.png)
``` powershell
overlayIcon = @{
    enable = $true
    size = 20.0
    lineWidth = 1.0  
    backgroundColor = "DeepPink"
    textColor = "White"
}
```

You can turn on/off the overlay badge feature by setting `enable` to `$true`/`$false`. You can also change the badge color. Available WPF color names are listed here:
https://docs.microsoft.com/en-us/dotnet/api/system.windows.media.colors?view=net-5.0

## Today's items summary

![items_summary](./docs/items_summary.png)

``` powershell
todaysRemainingItemsSummary = @{
    enable = $true
    maxItemCount = 6
    maxItemCharacterCount = 32
}
```

The summary of today's remaining events is displayed on top of the thumbnail window.

## Progress indicator

![progress_indicator](./docs/progress_indicator.png)

``` powershell
progressIndicator = @{
    enable = $true
    showProgressMinutesBefore = 30
}
```

When the next event is close, it is indicated by a progress indicator. The icon starts showing the progress indicator the minitues before the next event that is specified by `showProgressMinutesBefore`. When you are in the event, the bar is displayed in yellow.

## Click action

``` powershell
clickAction = @("FocusOnCalendar")
```

When the taskbar icon is clicked, the action you specify here is executed. The following actions are available:

|Action Name|Description|
|---|---|
|FocusOnCalendar|Opens the calendar view in Outlook keeping the previous view mode.|
|FocusOnToday|Opens the calendar view in Outlook and sets the view mode to Day.|
|FocusOnThisWeek|Opens the calendar view in Outlook and sets the view mode to Week.|
|FocusOnThisWorkWeek|Opens the calendar view in Outlook and sets the view mode to WorkWeek.|
|FocusOnThisMonth|Opens the calendar view in Outlook and sets the view mode to Month.|
|FocusOnNextNDays|Opens the calendar view in Outlook and sets the range to the number of days specified by the second argument. The number can be set to a value between 2 and 14.|
|OpenTodaysNextItem|Opens today's next item.|
|CreateNewAppointment|Opens a dialog to create a new appointment.|
|CreateNewMeeting|Opens a dialog to create a new meeting.|

## Thumb buttons

<img src="./docs/thumb_buttons.png" width=260>

``` powershell
thumbButtons = @(
    ,@{
            description = "Month"
            iconPath = "..\icons\month.png"
            clickAction = @("FocusOnThisMonth")
    }
)
```
You can add maximum 7 thumb buttons and assign actions performed when they are clicked. The formats of `iconPath` and `clickAction` are the same as the ones in the global settings.

# Tools

## [tools/list_outlook_calendar_folders.bat](./tools/list_outlook_calendar_folders.bat)

Lists all the Outlook calendar folder paths that the app can monitor. Copy one of the folder paths and paste it in your settings file.

## [tools/create_shortcut.bat](./tools/create_shortcut.bat)

This tool takes a settings file and creates a shortcut to run the app. If you want to create another app that monitors another Outlook calendar, you just need to create a settings file and run this tool again.

## [tools/convert_image_to_ico.bat](./tools/convert_image_to_ico.bat)

Converts an image to `.ico` file. When you want to update the icon of the shortcut, run this tool.

# Sample Icons

The icons except [icon.png](./icons/icon.png) were downloaded from [Google Material Icons](https://fonts.google.com/icons) which are distributed under [Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0.html).
