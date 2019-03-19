DEL /S /F /Q "%ALLUSERSPROFILE%\Application Data\Microsoft\Group Policy\History\*.*"
REG DELETE HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies /f
REG DELETE HKLM\SOFTWARE\Policies\Microsoft /f
REG DELETE HKLM\SOFTWARE\Policies\Microsoft /f
REG DELETE "HKCU\Software\Microsoft\Windows\Currentversion\Group Policy Objects" /f
REG DELETE HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies /f
DEL /F /Q C:\WINDOWS\security\Database\secedit.sdb
Klist purge
gpupdate /force /boot