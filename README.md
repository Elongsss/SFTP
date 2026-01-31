# SFTP
Script for copying json files to the linux using SFTP (WinSCP) and making local copy

This script writing for publicate json file with wallet course on the site our bank.

!Before use script, you must download WinSCP.exe and WinSCP.dll in directory, where you run this script!

# Credentials
```
$credentials = Get-Credential | Export-Clixml -Path 'C:\example\path\pass.xml'
```