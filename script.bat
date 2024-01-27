@echo off
setlocal

:: Define registry paths and keys
set "RegKey=HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
set "AppsUseLightThemeKey=AppsUseLightTheme"
set "SystemUsesLightThemeKey=SystemUsesLightTheme"

:: Read the current theme settings
for /f "tokens=3" %%a in ('reg query %RegKey% /v %AppsUseLightThemeKey%') do set /a "AppsUseLightTheme=%%a"
for /f "tokens=3" %%a in ('reg query %RegKey% /v %SystemUsesLightThemeKey%') do set /a "SystemUsesLightTheme=%%a"

:: Toggle the theme
if %AppsUseLightTheme%==1 (
    reg add %RegKey% /v %AppsUseLightThemeKey% /t REG_DWORD /d 0 /f
    reg add %RegKey% /v %SystemUsesLightThemeKey% /t REG_DWORD /d 0 /f
    echo Switched to Dark Mode.
) else (
    reg add %RegKey% /v %AppsUseLightThemeKey% /t REG_DWORD /d 1 /f
    reg add %RegKey% /v %SystemUsesLightThemeKey% /t REG_DWORD /d 1 /f
    echo Switched to Light Mode.
)

:: Use PowerShell to refresh the environment
powershell -command "Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class Settings { [DllImport(\"user32.dll\", SetLastError = true)] public static extern bool SendNotifyMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam); }' -PassThru; [Settings]::SendNotifyMessage((New-Object IntPtr -1), 0x1A, (New-Object IntPtr 0), (New-Object IntPtr 0))"

endlocal
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.SendKeys "{F5}"

 Set WshShell = CreateObject("WScript.Shell")
WshShell.AppActivate " - File Explorer"
WScript.Sleep 500
WshShell.SendKeys "{F5}"

 
