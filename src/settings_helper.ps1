function InitSettings($settingsPath, $scriptPath)
{
    . $settingsPath

    $dir = Split-Path $settingsPath -Parent
    $settings.path = $settingsPath
    $settings.directory = $dir
    $settings.scriptPath = $scriptPath
    $settings
}

function GetFullPathFromSettingsRelativePath($settings, $path)
{
    if (-not $path)
    {
        return ""
    }

    Push-Location $settings.directory
    $fullPath = Resolve-Path $path
    Pop-Location
    $fullPath.Path
}