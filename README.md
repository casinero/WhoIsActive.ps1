# WhoIsActive.ps1
PowerShell script to run sp_whoisactive and export a report.

Schedule it with `PowerShell.exe -ExecutionPolicy RemoteSigned -File WhoIsActive.ps1` in Windows Task Scheduler, or maybe as a SQL Server Agent Job triggered by Performance Condition Alert of `Processes blocked` counter?
