@echo off
echo ========================================
echo Jupiter Add-in Cache Clearing Script
echo ========================================
echo.

echo 1. Closing Word processes...
taskkill /f /im WINWORD.exe 2>nul
if %errorlevel%==0 (
    echo    Word processes closed.
) else (
    echo    No Word processes found.
)
echo.

echo 2. Clearing Office Add-in cache...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" (
    rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" 2>nul
    echo    Cleared: %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
)

if exist "%LOCALAPPDATA%\Microsoft\Office\Wef" (
    rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\Wef" 2>nul
    echo    Cleared: %LOCALAPPDATA%\Microsoft\Office\Wef
)

if exist "%APPDATA%\Microsoft\Office\16.0\Wef" (
    rmdir /s /q "%APPDATA%\Microsoft\Office\16.0\Wef" 2>nul
    echo    Cleared: %APPDATA%\Microsoft\Office\16.0\Wef
)

if exist "%APPDATA%\Microsoft\Office\Wef" (
    rmdir /s /q "%APPDATA%\Microsoft\Office\Wef" 2>nul
    echo    Cleared: %APPDATA%\Microsoft\Office\Wef
)
echo.

echo 3. Cache clearing completed!
echo.
echo NEXT STEPS:
echo -----------
echo 1. Open Visual Studio
echo 2. Open JuptiarAddins.sln
echo 3. Right-click "JuptiarAddins" project
echo 4. Select "Deploy" or "Rebuild"
echo 5. Start Word and check Jupiter tab
echo.
echo Expected changes:
echo - Three groups: Documents, Properties, Settings
echo - New icons for each button
echo - Better spacing and layout
echo.
pause
