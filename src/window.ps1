# process.ps1 must be included before executing this file.
# "Add-Type -AssemblyName PresentationFramework" must be called before including this file.

class Window
{
    $window
    $settings
    $timer
    $onClicked
    $lastOverlayCount = 0

    [void] Init($xamlPath, $title, $settings)
    {
        $this.settings = $settings

        $xaml = [xml](Get-Content $xamlPath)
        $nodeReader = (New-Object System.Xml.XmlNodeReader $xaml)
        $this.window = [System.Windows.Markup.XamlReader]::Load($nodeReader)
        $this.window.Title = $title

        $iconPath = GetFullPathFromSettingsRelativePath $settings $settings.iconPath
        if ($iconPath)
        {
            $this.window.Icon = $iconPath
        }

        # Start with Normal window to make Windows draw preview window.
        $this.window.WindowState = [System.Windows.WindowState]::Normal

        $class = $this

        $this.window.add_ContentRendered({
            $class.OnContentRendered()
        }.GetNewClosure())

        $this.window.add_StateChanged({
            $class.OnStateChanged()
        }.GetNewClosure())
    }

    [void] Term()
    {
    }

    [void] InitJumpList()
    {
        $jumpList = New-Object System.Windows.Shell.JumpList

        $jumpTask = New-Object System.Windows.Shell.JumpTask
        $jumpTask.Title = "Open settings location"
        $jumpTask.ApplicationPath = "explorer.exe"
        $jumpTask.IconResourcePath = "explorer.exe"
        $jumpTask.Arguments = "/select,`"{0}`"" -f $this.settings.path
        $jumpList.JumpItems.Add($jumpTask)

        $jumpTask = New-Object System.Windows.Shell.JumpTask
        $jumpTask.Title = "Run with debug console"
        $jumpTask.Arguments = "-ExecutionPolicy Bypass -File `"{0}`" `"{1}`" -SkipJumpList" -f $this.settings.scriptPath, $this.settings.path
        $jumpList.JumpItems.Add($jumpTask)

        $jumpList.Apply()
    }

    [void] OnContentRendered()
    {
        # Immediately minimize the window after the thumbnail is rendered.
        $this.window.WindowState = [System.Windows.WindowState]::Minimized
    }

    [void] OnStateChanged()
    {
        if ($this.window.WindowState -eq [System.Windows.WindowState]::Minimized)
        {
            return
        }

        if ($this.onClicked)
        {
            $this.onClicked.Invoke()
        }

        $this.window.WindowState = [System.Windows.WindowState]::Minimized
    }

    [void] SetOnClickedFunction($block)
    {
        $this.onClicked = $block
    }

    [void] SetTaskbarItemInfoDescription($text)
    {
        $this.window.TaskbarItemInfo.Description = $text
    }

    [Object] AddThumbButton($thumbButtonSetting)
    {
        $button = New-Object System.Windows.Shell.ThumbButtonInfo
        $button.Description = $thumbButtonSetting.description
        $button.DismissWhenClicked = $true

        $iconPath = GetFullPathFromSettingsRelativePath $this.settings $thumbButtonSetting.iconPath
        if ($iconPath)
        {
            $button.ImageSource = $iconPath
        }

        $this.window.TaskbarItemInfo.ThumbButtonInfos.Add($button)
        return $button
    }

    [void] SetProgressIndicator($progress)
    {
        $progress = [Math]::Max([Math]::Min($progress, 1), 0)
        if ($progress -eq 0)
        {
            $state = [System.Windows.Shell.TaskbarItemProgressState]::None
        }
        elseif ($progress -eq 1)
        {
            $state = [System.Windows.Shell.TaskbarItemProgressState]::Paused
        }
        else
        {
            $state = [System.Windows.Shell.TaskbarItemProgressState]::Normal
        }
        $this.window.TaskbarItemInfo.ProgressState = $state
        $this.window.TaskbarItemInfo.ProgressValue = $progress
    }

    [void] ShowDialog()
    {
        $this.window.ShowDialog()
    }

    [void] UpdateOverlayCount($count)
    {
        if ($count -eq $this.lastOverlayCount)
        {
            return
        }

        $this.lastOverlayCount = $count

        if ($count -eq 0)
        {
            $content = ""
        }
        elseif ($count -lt 0)
        {
            $content = "E"
        }
        else
        {
            $content = [Math]::Min($count, 99).ToString()
        }
        $this.UpdateOverlayIcon($content)
    }

    [void] UpdateOverlayIcon($content)
    {
        if (-not $content)
        {
            $this.window.TaskbarItemInfo.Overlay = $null
            return
        }
        
        $dpi = 96
        $iconParameters = $this.GetOverlayIconParameters()
        $iconParameters.Text = $content

        $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($iconParameters.IconSize, $iconParameters.IconSize, $dpi, $dpi, [System.Windows.Media.PixelFormats]::Default)
        $rect = New-Object System.Windows.Rect 0, 0, $iconParameters.IconSize, $iconParameters.IconSize
        $control = New-Object System.Windows.Controls.ContentControl
        $control.ContentTemplate = $this.window.Resources["OverlayIcon"]
        $control.Content = [PSCustomObject]$iconParameters
        $control.Arrange($rect)
        $bitmap.Render($control)
        $this.window.TaskbarItemInfo.Overlay = $bitmap
    }

    [Object] GetOverlayIconParameters()
    {
        $parameters = @{}
        $parameters.IconSize = 20.0
        if ($this.settings.overlayIcon.size)
        {
            $parameters.IconSize = $this.settings.overlayIcon.size
        }
        $parameters.FontSize = $parameters.IconSize * 0.7

        $parameters.LineWidth = 1.0
        if ($this.settings.overlayIcon.lineWidth)
        {
            $parameters.LineWidth = $this.settings.overlayIcon.lineWidth
        }

        $parameters.Color = $this.settings.overlayIcon.backgroundColor
        $parameters.TextColor = $this.settings.overlayIcon.textColor

        return $parameters
    }

    [void] StartTimerFunction($block, $intervalInSeconds)
    {
        if ($this.timer)
        {
            $this.timer.Stop()
        }
        $this.timer = New-Object System.Windows.Threading.DispatcherTimer
        $this.timer.interval = New-Object TimeSpan(0, 0, $intervalInSeconds)
        $this.timer.add_tick($block)
        $this.timer.Start()
    }
}
