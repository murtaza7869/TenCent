Invoke-WebRequest 'https://dl.dropboxusercontent.com/s/jdagxkbgubqilq2/setup.exe' -OutFile 'C:\Windows\temp\setup.exe'
Invoke-WebRequest 'https://dl.dropboxusercontent.com/s/1yy4rzlky2nqz5s/configuration-Office365-x64.xml' -OutFile 'C:\Windows\temp\configuration-Office365-x64.xml'
$arg="/configure C:\Windows\temp\configuration-Office365-x64.xml"
Start-Process -FilePath "C:\Windows\temp\setup.exe" $arg -WorkingDirectory "C:\Windows\temp"
