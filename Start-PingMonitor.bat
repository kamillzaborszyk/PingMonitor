@echo off
:: PingMonitor Launcher — double-click to start
:: Hides the PowerShell console window automatically
start "" powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0PingMonitor.ps1"
