﻿# The encoding of this file must be UTF-8 with BOM.

$settings = @{

    outlook = @{
        # The outlook calender folder full path.
        folderPath = "\\your-email-address@sample.com\calendar-folder-name"
    }

    # Icon file used for the title bar. The path should be either a relative path from this settings file or a full path.
    iconPath = "..\icons\icon.png"

    # Today's item count is queried with this interval.
    updateIntervalInSeconds = 3

    # Overlay badge icon settings
    overlayIcon = @{
        # $true or $false.
        enable = $true

        # Available WPF color names are listed here:
        # https://docs.microsoft.com/en-us/dotnet/api/system.windows.media.colors?view=net-5.0
        backgroundColor = "LightCoral"
        textColor = "White"
    }

    clickAction = "FocusOnCalendar"

    # Thumb buttons
    # You can add max 7 buttons.
    thumbButtons = @(
        ,@{
            description = "Today"
            iconPath = "..\icons\day.png"
            clickAction = "FocusOnToday"
        }
        ,@{
            description = "Week"
            iconPath = "..\icons\week.png"
            clickAction = "FocusOnThisWeek"
        }
        ,@{
            description = "Month"
            iconPath = "..\icons\month.png"
            clickAction = "FocusOnThisMonth"
        }
    )
}