@ECHO OFF
PowerShell.exe -Command "& {Start-Process PowerShell.exe -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0GroupManager.ps1""' -Verb RunAs}"