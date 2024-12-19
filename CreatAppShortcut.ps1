# Define the path to the application .exe file and the shortcut name
$exePath = "C:\Path\To\YourApplication.exe"  # Replace with the actual path to your .exe file
$shortcutName = "Launch Your Application"

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
Create-Shortcut -targetPath $exePath -shortcutPath $publicDesktopPath -shortcutName $shortcutName

# Create the Start Menu Shortcut
Create-Shortcut -targetPath $exePath -shortcutPath $startMenuPath -shortcutName $shortcutName

# Pin to Taskbar using explorer.exe
$desktopShortcutPath = [System.IO.Path]::Combine($publicDesktopPath, "$shortcutName.lnk")
if (Test-Path -Path $desktopShortcutPath) {
    # Launch the explorer command to pin to the taskbar
    Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$desktopShortcutPath`""
    Start-Sleep -Seconds 2
    [void][System.Windows.Forms.SendKeys]::SendWait("+{F10}p")  # Simulates Shift+F10 and Pin to Taskbar
}

Write-Output "Shortcut created successfully on Desktop, Start Menu, and Taskbar."
