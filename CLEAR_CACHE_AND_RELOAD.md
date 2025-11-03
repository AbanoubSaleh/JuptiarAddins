# Clear Office Cache and Reload Settings Icon

## Summary of Changes Made

1. ✅ Updated manifest to reference `Settings16x16.png`, `Settings32x32.png`, `Settings80x80.png`
2. ✅ Replaced old settings icon files with new gear icon
3. ✅ Incremented manifest version from 1.4.1.0 to 1.4.2.0
4. ✅ Rebuilt the project

## The Problem

Office/Word aggressively caches add-in manifests and icons. Even though we've updated the files, Word is still showing the old cached icons.

## Solution: Clear Office Cache

### Method 1: Clear Office Cache Folders (RECOMMENDED)

1. **Close Word completely** (make sure no WINWORD.EXE processes are running)

2. **Delete Office cache folders**:
   - Open File Explorer and paste each path below into the address bar
   - Delete all contents in each folder:

   ```
   %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
   ```
   
   ```
   %LOCALAPPDATA%\Microsoft\Office\Wef
   ```
   
   ```
   %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\INetCache
   ```

3. **Restart Word**

4. The new gear icon should now appear!

### Method 2: Remove and Re-add the Add-in

1. Open Word
2. Go to **File > Options > Add-ins**
3. At the bottom, select **Manage: COM Add-ins** and click **Go**
4. **Uncheck** "Jupiter Document Manager" and click **OK**
5. Go back to **File > Options > Add-ins**
6. Select **Manage: COM Add-ins** and click **Go** again
7. **Check** "Jupiter Document Manager" and click **OK**

### Method 3: Run from Visual Studio (Forces Fresh Load)

1. Open the solution in Visual Studio
2. Press **F5** to run/debug the add-in
3. This will force Office to reload everything fresh

## Verification

After clearing the cache, the Settings button in the Jupiter ribbon should show the new gear icon instead of the old settings icon.

## Technical Details

- **Icon files location**: `JuptiarAddinsWeb\Images\`
- **Manifest location**: `JuptiarAddins\JuptiarAddinsManifest\JuptiarAddins.xml`
- **Compiled manifest**: `JuptiarAddins\bin\Debug\OfficeAppManifests\JuptiarAddinsManifest.xml`
- **Web server**: IIS Express on `https://localhost:44316`
- **Icon URLs**: 
  - `https://localhost:44316/Images/Settings16x16.png`
  - `https://localhost:44316/Images/Settings32x32.png`
  - `https://localhost:44316/Images/Settings80x80.png`

