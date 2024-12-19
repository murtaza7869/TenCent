# Define the path to the application .exe file and the shortcut name
$exePath = "C:\files\SPYXX.EXE"  # Replace with the actual path to your .exe file
$shortcutName = "ISPY"

# Paths for Public Desktop, Start Menu, and Taskbar
$publicDesktopPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("CommonDesktopDirectory"))
$startMenuPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("CommonStartMenu"), "Programs")
$taskbarPinnerPath = [Environment]::ExpandEnvironmentVariables("%AppData%\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")

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
Create-Shortcut -targetPath $exePath -shortcutPath $publicDesktopPath -shortcutName $shortcutName

# Create the Start Menu Shortcut
Create-Shortcut -targetPath $exePath -shortcutPath $startMenuPath -shortcutName $shortcutName

# Pin to Taskbar (requires the shortcut to exist first)
$taskbarShortcutPath = [System.IO.Path]::Combine($taskbarPinnerPath, "$shortcutName.lnk")
if (!(Test-Path -Path $taskbarShortcutPath)) {
    Copy-Item -Path ([System.IO.Path]::Combine($publicDesktopPath, "$shortcutName.lnk")) -Destination $taskbarShortcutPath -Force
    Invoke-Expression "explorer shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}"  # Refresh the Taskbar
}

Write-Output "Shortcut created successfully on Desktop, Start Menu, and Taskbar."
