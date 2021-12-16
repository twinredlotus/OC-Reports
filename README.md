# OC-Reports
Email compressed archive of weekly report files.
## Run command against remote machines with psremoting enabled.

Example:
--------------------------------------------------
    powershell.exe -ep bypass -command { Invoke-Command -ComputerName name -Credential $cred -FilePath path\to\fileonpc }