Invoke-WebRequest 'https://dl.dropboxusercontent.com/s/jdagxkbgubqilq2/setup.exe' -OutFile 'C:\Windows\temp\setup.exe'
Invoke-WebRequest 'https://github.com/murtaza7869/TenCent/raw/main/TencentOfficeConfig.xml' -OutFile 'C:\Windows\temp\TencentOfficeConfig.xml'
$arg="/configure C:\Windows\temp\TencentOfficeConfig.xml"
Start-Process -FilePath "C:\Windows\temp\setup.exe" $arg -WorkingDirectory "C:\Windows\temp"
