@echo off
cls
echo Enable "File System Audit"
auditpol /set /subcategory:{0CCE921D-69AE-11D9-BED3-505054503030} /success:enable
echo ------------------------------------------------------------
auditpol /get /subcategory:{0CCE921D-69AE-11D9-BED3-505054503030}

powershell .\enable_voice_folder_audit.ps1

