# Parse script arguments
param (
    [Parameter(Mandatory = $true)]
    [string]$ExePath,   # Path to the application .exe file
    [Parameter(Mandatory = $true)]
    [string]$ShortcutName  # Name of the shortcut
)

# Paths for Public Desktop and Start Menu
$publicDesktopPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("CommonDesktopDirectory"))
$startMenuPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("CommonStartMenu"), "Programs")

# Ensure Start Menu folder exists
if (!(Test-Path -Path $startMenuPath)) {
    New-Item -ItemType Directory -Path $startMenuPath -Force
}

# Create a WScript.Shell COM object
$wScriptShell = New-Object -ComObject WScript.Shell

# Function to create a shortcut
function Create-Shortcut {
    param (
        [string]$targetPath,
        [string]$shortcutPath,
        [string]$shortcutName
    )
    $shortcut = $wScriptShell.CreateShortcut([System.IO.Path]::Combine($shortcutPath, "$shortcutName.lnk"))
    $shortcut.TargetPath = $targetPath
    $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetPath)
    $shortcut.IconLocation = $targetPath
    $shortcut.Save()
}

# Create the Desktop Shortcut
Create-Shortcut -targetPath $ExePath -shortcutPath $publicDesktopPath -shortcutName $ShortcutName

# Create the Start Menu Shortcut
Create-Shortcut -targetPath $ExePath -shortcutPath $startMenuPath -shortcutName $ShortcutName

# Pin to Taskbar using custom PowerShell function
function Pin-ToTaskbar {
    param (
        [string]$shortcutPath
    )
    $verb = "pin to taskbar"
    $shellApp = New-Object -ComObject Shell.Application
    $file = $shellApp.Namespace((Get-Item $shortcutPath).DirectoryName).ParseName((Get-Item $shortcutPath).Name)

    foreach ($v in $file.Verbs()) {
        if ($v.Name -replace '&', '' -match $verb) {
            $v.DoIt()
            Write-Output "Pinned to Taskbar: $shortcutPath"
            return
        }
    }
    Write-Warning "Pin to Taskbar option not found for $shortcutPath"
}

# Pin the desktop shortcut to Taskbar
$desktopShortcutPath = [System.IO.Path]::Combine($publicDesktopPath, "$ShortcutName.lnk")
Pin-ToTaskbar -shortcutPath $desktopShortcutPath

Write-Output "Shortcut created successfully on Desktop, Start Menu, and Taskbar."
