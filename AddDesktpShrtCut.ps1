# Define the URL and the shortcut name
$url = "https://wbte.drcedirect.com/TABE/#portal/tabe/595219/adminId=59521"
$shortcutName = "Launch TABE Portal"

# Get the path to the 'Public Desktop' where shortcuts appear for all users
$publicDesktopPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("CommonDesktopDirectory"))

# Path to the shortcut file
$shortcutPath = [System.IO.Path]::Combine($publicDesktopPath, "$shortcutName.url")

# Create the shortcut file
$shortcutContent = @"
[InternetShortcut]
URL=$url
IconFile=%SystemRoot%\System32\SHELL32.dll
IconIndex=13
"@

# Write the shortcut to the file
Set-Content -Path $shortcutPath -Value $shortcutContent -Encoding ASCII

Write-Output "Shortcut created successfully at $shortcutPath"
