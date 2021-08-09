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
        
        $width = 20
        $height = 20
        $dpi = 96
        $backgroundColor = $this.settings.overlayIcon.backgroundColor
        $textColor = $this.settings.overlayIcon.textColor

        $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($width, $height, $dpi, $dpi, [System.Windows.Media.PixelFormats]::Default)
        $rect = New-Object System.Windows.Rect 0, 0, $width, $height
        $control = New-Object System.Windows.Controls.ContentControl
        $control.ContentTemplate = $this.window.Resources["OverlayIcon"]
        $control.content = [PSCustomObject]@{
            Color = $backgroundColor
            TextColor = $textColor
            Text = $content
        }
        $control.Arrange($rect)
        $bitmap.Render($control)
        $this.window.TaskbarItemInfo.Overlay = $bitmap
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
