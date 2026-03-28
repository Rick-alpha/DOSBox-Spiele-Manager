DOSBox Game Manager - UI Description
===========================================

Short Description
translated with Google
----------------

Requires the original DOSBox: https://www.dosbox.com/ A great piece of software, support the guys!

The DOSBox Game Manager is a Windows interface (PowerShell + WinForms)
that lets you manage your DOS game library in a main folder.

The UI helps you scan, sort, launch, and optimize your games
as well as manage thumbnails and favorites.

What the UI does

----------------
1. It remembers your last used game folder.

2. It scans subfolders for games and determines the following for each entry:

- Game name/folder

- Possible launch file (.exe/.com/.bat)

- Optimization status (e.g., existing .conf file or DOSBox call in .bat)

3. It displays all games as a list or tile view.

4. It provides detailed information and actions for the selected game.

5. It can cache data and preview images so that everything doesn't have to be reloaded the next time it's launched.
... UI Areas at a Glance
-------------------------
Header (top):

- Folder Path (text field): Directly enter game folder

- Select Folder: Open folder dialog

- Rescan: Manually start scan

- View: Toggle list/tiles

- Save Data: Save entry data to cache

- Load Data: Load entry data from cache

- Create Previews: Create/update preview images for all loaded games

- Search + X: Enter filter text and quickly delete

- Show Favorites Only / Show All: Toggle favorites filter

- Sort: A-Z, Z-A, Status, Favorites

- Hide/Show Header: More space in the workspace

Left Side (Library):

- Game view as a table or tile

- Preview images are displayed in tile view, if available

Right Side (Details + Actions):

- Detail fields for name, folder, startup file, optimization, detection, config

- Set/remove favorite
- Game Start

- Optimize (create configuration file and launcher)

- Customize (config) in Notepad

- Open DOSBox options

- Add images (local)

- Google image (one game)

- Google images (ALL) for batch download

- Log window with status and error messages

Important functions in detail

-----------------------------
Start game:

- The manager prioritizes launching existing startup scripts directly.

- If necessary, DOSBox is launched with the appropriate configuration.

Optimize:

- Creates a basic DOSBox configuration in the game folder.

- Additionally creates a suitable launcher.

- Marks the entry as optimized afterward.

Preview images:

- Automatically created from existing image files in the game folder.

- Manually set from a local image file.

- Online search (e.g., cover/screenshot) for a single game or as a batch search for all games.

Favorites:

- Games can be marked as favorites.

- Favorites can be filtered and sorted by preference.

Cache and Persistence:

- Settings (last used root folder) are saved.

- Entries, including favorites/preview information, can be saved as a cache.

- The cache can be loaded the next time the program is launched.


Technical Details: Paths in the Code (2-Column Overview)
---------------------------------------------
Column 1: Area/Function

Column 2: Where in the code and which path logic is used

- Base folder of the manager

- -> DOSBox-Games-Manager.ps1: `$script:BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path`

- Meaning: All relative paths are built from the script's folder.

- DOSBox main program

- -> DOSBox-Games-Manager.ps1: `$script:DosBoxExe = Join-Path $script:BaseDir 'DOSBox.exe'`

- Required path: `DOSBox.exe` must be located in the same main folder as the manager.



```` - Options Shortcut/Launcher (technically BAT)

-> DOSBox-Game-Manager.ps1: `$script:OptionsBat = Join-Path $script:BaseDir 'DOSBox 0.74-3 Options.bat'`

-> The `DOSBox Options` UI button launches this exact BAT file.

- Saved Last Game Folder

-> DOSBox-Game-Manager.ps1: `$script:SettingsPath = Join-Path $script:BaseDir 'dosbox-game-manager.settings.json'`

-> The last used root folder is stored here.


- Cache file per game root

- -> DOSBox-Games-Manager.ps1: `Get-CachePath([string]$rootFolder)`

- -> Result: `<Game-Root>\\dosbox-manager-cache.json`

- Preview image per game folder

- -> DOSBox-Games-Manager.ps1: `Get-PreviewCachePath([string]$folderPath)`

- -> Result: `<Game-Folder>\\.dosbox-preview.png`

- Hard startup check for DOSBox

- -> DOSBox-Games-Manager.ps1: `if (-not (Test-Path $script:DosBoxExe)) { ... exit 1 }`

- -> If `DOSBox.exe` is missing, the manager will not start.



``` Important information about shortcuts (.lnk)

-------------------------------- The UI internally uses actual file paths to `DOSBox.exe` and `.bat/.conf`.

`.lnk` files in the package are convenient shortcuts for Windows, but not the
prima
