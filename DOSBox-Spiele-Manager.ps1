Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'

$script:BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:DosBoxExe = Join-Path $script:BaseDir 'DOSBox.exe'
$script:OptionsBat = Join-Path $script:BaseDir 'DOSBox 0.74-3 Options.bat'
$script:SettingsPath = Join-Path $script:BaseDir 'dosbox-game-manager.settings.json'
$script:FolderCacheFileName = 'dosbox-manager-cache.json'
$script:PreviewCacheFileName = '.dosbox-preview.png'
$script:CurrentRoot = ''
$script:Entries = @()
$script:IsScanning = $false
$script:PreviewsEnabled = $false
$script:AutoScanOnStartup = $false
$script:SortMode = 'name-asc'
$script:ShowFavoritesOnly = $false
$script:FilterText = ''
$script:UmlautU = [char]0x00DC
$script:umlautu = [char]0x00FC

function Enable-WebTls {
    try {
        $all = [Net.SecurityProtocolType]::Tls12
        try { $all = $all -bor [Net.SecurityProtocolType]::Tls11 } catch {}
        try { $all = $all -bor [Net.SecurityProtocolType]::Tls } catch {}
        [Net.ServicePointManager]::SecurityProtocol = $all
    } catch {
        # Keep defaults on very old environments.
    }
}

function Show-Error([string]$message) {
    [System.Windows.Forms.MessageBox]::Show($message, 'DOSBox Spiele-Manager', 'OK', 'Error') | Out-Null
}

function Load-Settings {
    if (Test-Path $script:SettingsPath) {
        try {
            $settings = Get-Content -Path $script:SettingsPath -Raw | ConvertFrom-Json
            if ($settings.RootFolder -and (Test-Path $settings.RootFolder)) {
                return $settings.RootFolder
            }
        } catch {
            # Ignore malformed settings.
        }
    }
    return ''
}

function Save-Settings([string]$rootFolder) {
    $payload = @{ RootFolder = $rootFolder } | ConvertTo-Json
    Set-Content -Path $script:SettingsPath -Value $payload -Encoding UTF8
}

function Get-CachePath([string]$rootFolder) {
    return (Join-Path $rootFolder $script:FolderCacheFileName)
}

function Save-FolderCache([string]$rootFolder, [object[]]$entries) {
    if (-not (Test-Path $rootFolder)) {
        return $false
    }

    $cachePath = Get-CachePath -rootFolder $rootFolder
    $payload = [PSCustomObject]@{
        RootFolder = $rootFolder
        SavedAt = (Get-Date).ToString('s')
        Entries = @($entries | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                FolderPath = $_.FolderPath
                RootFolder = $_.RootFolder
                RelativePath = $_.RelativePath
                ExecutableName = $_.ExecutableName
                ExecutableFullPath = $_.ExecutableFullPath
                IsOptimized = [bool]$_.IsOptimized
                ConfigPath = $_.ConfigPath
                OptimizationReason = $_.OptimizationReason
                OptimizationSource = if ($_.OptimizationSource) { $_.OptimizationSource } else { 'existing' }
                PreviewImagePath = if ($_.PreviewImagePath) { $_.PreviewImagePath } else { '' }
                PreviewSource = if ($_.PreviewSource) { $_.PreviewSource } else { 'none' }
                PreviewImageUrl = if ($_.PreviewImageUrl) { $_.PreviewImageUrl } else { '' }
                PreviewUpdatedAt = if ($_.PreviewUpdatedAt) { $_.PreviewUpdatedAt } else { '' }
                IsFavorite = [bool]$_.IsFavorite
            }
        })
    }

    $json = $payload | ConvertTo-Json -Depth 6
    Set-Content -Path $cachePath -Value $json -Encoding UTF8
    return $true
}

function Load-FolderCache([string]$rootFolder) {
    $cachePath = Get-CachePath -rootFolder $rootFolder
    if (-not (Test-Path $cachePath)) {
        return $null
    }

    try {
        $parsed = Get-Content -Path $cachePath -Raw | ConvertFrom-Json
        if (-not $parsed.Entries) {
            return @()
        }

        $loaded = @($parsed.Entries | ForEach-Object {
            $fallbackPreviewPath = ''
            if ($_.FolderPath -and (Test-Path $_.FolderPath)) {
                $candidate = Get-PreviewCachePath -folderPath $_.FolderPath
                if (Test-Path $candidate) {
                    $fallbackPreviewPath = $candidate
                }
            }

            [PSCustomObject]@{
                Name = $_.Name
                FolderPath = $_.FolderPath
                RootFolder = if ($_.RootFolder) { $_.RootFolder } else { $rootFolder }
                RelativePath = $_.RelativePath
                ExecutableName = $_.ExecutableName
                ExecutableFullPath = $_.ExecutableFullPath
                IsOptimized = [bool]$_.IsOptimized
                ConfigPath = $_.ConfigPath
                OptimizationReason = $_.OptimizationReason
                OptimizationSource = if ($_.OptimizationSource) { $_.OptimizationSource } else { 'existing' }
                PreviewImagePath = if ($_.PreviewImagePath) { $_.PreviewImagePath } elseif ($fallbackPreviewPath) { $fallbackPreviewPath } else { '' }
                PreviewSource = if ($_.PreviewSource) { $_.PreviewSource } elseif ($fallbackPreviewPath) { 'cached' } else { 'none' }
                PreviewImageUrl = if ($_.PreviewImageUrl) { $_.PreviewImageUrl } else { '' }
                PreviewUpdatedAt = if ($_.PreviewUpdatedAt) { $_.PreviewUpdatedAt } else { '' }
                IsFavorite = if ($_.IsFavorite) { [bool]$_.IsFavorite } else { $false }
            }
        })

        return $loaded
    } catch {
        return $null
    }
}

function Get-StringHash([string]$text) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    $md5 = [System.Security.Cryptography.MD5]::Create()
    try {
        $hash = $md5.ComputeHash($bytes)
        return ([BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    } finally {
        $md5.Dispose()
    }
}

function Get-PreviewCachePath([string]$folderPath) {
    return (Join-Path $folderPath $script:PreviewCacheFileName)
}

function Get-PreviewSourceFile([string]$folderPath) {
    $priorityPatterns = @(
        'cover*', 'title*', 'screenshot*', 'shot*', 'preview*', 'menu*'
    )

    $allowed = @('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.pcx', '.tga', '.lbm', '.iff', '.pic')
    $files = @(Get-ChildItem -Path $folderPath -File -ErrorAction SilentlyContinue)

    foreach ($pattern in $priorityPatterns) {
        $match = $files |
            Where-Object { $_.BaseName -like $pattern } |
            Where-Object { $allowed -contains $_.Extension.ToLowerInvariant() } |
            Sort-Object Name |
            Select-Object -First 1
        if ($match) {
            return $match
        }
    }

    $fallback = $files |
        Where-Object { $allowed -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object Name |
        Select-Object -First 1

    return $fallback
}

function Convert-ImageToPreview([string]$sourceImagePath, [string]$outputImagePath) {
    $stream = $null
    $image = $null
    $thumb = $null
    $graphics = $null
    try {
        $stream = [System.IO.File]::OpenRead($sourceImagePath)
        $image = [System.Drawing.Image]::FromStream($stream)

        $targetWidth = 160
        $targetHeight = 120
        $thumb = New-Object System.Drawing.Bitmap($targetWidth, $targetHeight)
        $graphics = [System.Drawing.Graphics]::FromImage($thumb)
        $graphics.Clear([System.Drawing.Color]::Black)
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

        $scaleX = $targetWidth / [double]$image.Width
        $scaleY = $targetHeight / [double]$image.Height
        $scale = [Math]::Min($scaleX, $scaleY)

        $drawWidth = [int]([Math]::Round($image.Width * $scale))
        $drawHeight = [int]([Math]::Round($image.Height * $scale))
        $offsetX = [int](($targetWidth - $drawWidth) / 2)
        $offsetY = [int](($targetHeight - $drawHeight) / 2)

        $graphics.DrawImage($image, $offsetX, $offsetY, $drawWidth, $drawHeight)
        $thumb.Save($outputImagePath, [System.Drawing.Imaging.ImageFormat]::Png)
        return $true
    } catch {
        return $false
    } finally {
        if ($graphics) { $graphics.Dispose() }
        if ($thumb) { $thumb.Dispose() }
        if ($image) { $image.Dispose() }
        if ($stream) { $stream.Dispose() }
    }
}

function Ensure-EntryPreview(
    [pscustomobject]$entry,
    [System.Windows.Forms.ImageList]$imageList,
    [bool]$forceRebuild,
    [System.Windows.Forms.TextBox]$logBox
) {
    if (-not $entry -or -not $entry.FolderPath -or -not (Test-Path $entry.FolderPath)) {
        return 'default'
    }

    $key = Get-StringHash -text $entry.FolderPath
    if ($imageList.Images.ContainsKey($key)) {
        return $key
    }

    $previewPath = Get-PreviewCachePath -folderPath $entry.FolderPath
    if ($forceRebuild -or -not (Test-Path $previewPath)) {
        $source = Get-PreviewSourceFile -folderPath $entry.FolderPath
        if ($source) {
            $ok = Convert-ImageToPreview -sourceImagePath $source.FullName -outputImagePath $previewPath
            if ($ok -and $logBox) {
                $logBox.AppendText("Preview erstellt: $($entry.Name) <- $($source.Name)`r`n")
            }
            if ($ok) {
                $entry.PreviewImagePath = $previewPath
                if (-not $entry.PreviewSource -or $entry.PreviewSource -eq 'none') {
                    $entry.PreviewSource = 'auto'
                }
                if (-not $entry.PreviewUpdatedAt) {
                    $entry.PreviewUpdatedAt = (Get-Date).ToString('s')
                }
            }
        }
    }

    if (Test-Path $previewPath) {
        try {
            $bmp = New-Object System.Drawing.Bitmap($previewPath)
            $imageList.Images.Add($key, $bmp)
            $bmp.Dispose()
            $entry.PreviewImagePath = $previewPath
            if (-not $entry.PreviewSource) {
                $entry.PreviewSource = 'cached'
            }
            return $key
        } catch {
            return 'default'
        }
    }

    return 'default'
}

function Ensure-TileImageList([System.Windows.Forms.ListView]$listView) {
    if ($script:tileImageList) {
        return $script:tileImageList
    }

    try {
        $script:tileImageList = New-Object System.Windows.Forms.ImageList
        $script:tileImageList.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit
        $script:tileImageList.ImageSize = New-Object System.Drawing.Size(160, 120)

        $defaultPreview = New-Object System.Drawing.Bitmap(160, 120)
        $defaultGraphics = [System.Drawing.Graphics]::FromImage($defaultPreview)
        $defaultGraphics.Clear([System.Drawing.Color]::FromArgb(40, 40, 40))
        $defaultGraphics.Dispose()
        $script:tileImageList.Images.Add('default', $defaultPreview)

        $listView.LargeImageList = $script:tileImageList
        $listView.SmallImageList = $script:tileImageList
        return $script:tileImageList
    } catch {
        return $null
    }
}

function Decode-WebEscapedUrl([string]$rawUrl) {
    if ([string]::IsNullOrWhiteSpace($rawUrl)) {
        return $null
    }

    $decoded = $rawUrl
    $decoded = $decoded -replace '\\\\/', '/'
    $decoded = [regex]::Replace(
        $decoded,
        '\\u([0-9a-fA-F]{4})',
        {
            param($m)
            [char][int]::Parse($m.Groups[1].Value, [System.Globalization.NumberStyles]::HexNumber)
        }
    )

    try {
        $decoded = [System.Uri]::UnescapeDataString($decoded)
    } catch {
        # Keep best-effort decoded value.
    }

    return $decoded
}

function Is-UsableImageUrl([string]$url) {
    if ([string]::IsNullOrWhiteSpace($url)) {
        return $false
    }

    if ($url -notmatch '^https?://') {
        return $false
    }

    if ($url -match '(?i)(gstatic\.com|google\.com/images|data:image)') {
        return $false
    }

    if ($url -match '(?i)\.(jpg|jpeg|png|webp|bmp|gif)(\?|$)') {
        return $true
    }

    return $true
}

function Get-FirstGoogleImageUrl([string]$query) {
    if ([string]::IsNullOrWhiteSpace($query)) {
        return $null
    }

    try {
        Enable-WebTls
        $encoded = [System.Uri]::EscapeDataString($query)
        $searchUrl = "https://www.google.com/search?tbm=isch&q=$encoded"
        
        $resp = Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -TimeoutSec 10 `
            -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
        $html = [string]$resp.Content
        
        if ($html -match '"ou":"([^"]+?\.(jpg|jpeg|png|webp))"') {
            $url = $matches[1]
            if ($url -match '^https?://' -and $url -notmatch 'gstatic|google\.com') {
                return $url
            }
        }
    } catch {
        # Silent, continue to fallback
    }

    return $null
}

function Get-FirstBingImageUrl([string]$query) {
    if ([string]::IsNullOrWhiteSpace($query)) {
        return $null
    }

    try {
        Enable-WebTls
        $encoded = $query -replace ' ', '+'
        $searchUrl = "https://www.bing.com/images/search?q=$encoded"
        
        $resp = Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -TimeoutSec 10 `
            -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
        $html = [string]$resp.Content
        
        if ($html -match 'murl&quot;:&quot;([^&quot;]+?\.(jpg|jpeg|png|gif|webp))') {
            $url = $matches[1]
            if ($url -match '^https?://') {
                return $url
            }
        }
        
        if ($html -match '"murl":"([^"]+?\.(jpg|jpeg|png|gif|webp))"') {
            $url = $matches[1]
            if ($url -match '^https?://') {
                return $url
            }
        }
        
        if ($html -match 'src="([^"]*\.(jpg|jpeg|png|gif|webp)[^"]*)') {
            $url = $matches[1]
            if ($url -match '^https?://') {
                return $url
            }
        }
    } catch {
        # Silent, continue to fallback
    }

    return $null
}

function Get-FirstDuckDuckGoImageUrl([string]$query) {
    if ([string]::IsNullOrWhiteSpace($query)) {
        return $null
    }

    try {
        Enable-WebTls
        $encoded = $query -replace ' ', '+'
        $searchUrl = "https://duckduckgo.com/?q=$encoded+image&iax=images&ia=images"
        
        $resp = Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -TimeoutSec 10 `
            -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
        $html = [string]$resp.Content
        
        if ($html -match '"image":"([^"]+?\.(jpg|jpeg|png|gif|webp))"') {
            $url = $matches[1]
            if ($url -match '^https?://') {
                return $url
            }
        }
        
        if ($html -match 'img src="([^"]+?\.(jpg|jpeg|png|gif|webp))"') {
            $url = $matches[1]
            if ($url -match '^https?://') {
                return $url
            }
        }
    } catch {
        # Silent, continue to fallback
    }

    return $null
}

function Get-ImageFromArchiveOrg([string]$gameName) {
    if ([string]::IsNullOrWhiteSpace($gameName)) {
        return $null
    }

    try {
        Enable-WebTls
        $encoded = $gameName -replace ' ', '%20'
        $searchUrl = "https://archive.org/advancedsearch.php?q=title:$encoded%20AND%20mediatype:software&output=json&rows=5"
        
        $resp = Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -TimeoutSec 10
        $json = $resp.Content | ConvertFrom-Json -ErrorAction SilentlyContinue
        
        if ($json.response.docs -and $json.response.docs.Count -gt 0) {
            $firstDoc = $json.response.docs[0]
            $itemId = $firstDoc.identifier
            if ($itemId) {
                $thumbUrl = "https://archive.org/services/img/$itemId"
                if ($thumbUrl -match '^https?://') {
                    return $thumbUrl
                }
            }
        }
    } catch {
        # Silent, continue to fallback
    }

    return $null
}

function Get-FirstOnlineImageUrl([string]$query) {
    $googleUrl = Get-FirstGoogleImageUrl -query $query
    if ($googleUrl) {
        return @{
            Url = $googleUrl
            Source = 'google'
        }
    }

    $bingUrl = Get-FirstBingImageUrl -query $query
    if ($bingUrl) {
        return @{
            Url = $bingUrl
            Source = 'bing'
        }
    }

    $duckUrl = Get-FirstDuckDuckGoImageUrl -query $query
    if ($duckUrl) {
        return @{
            Url = $duckUrl
            Source = 'duckduckgo'
        }
    }

    $archiveUrl = Get-ImageFromArchiveOrg -gameName $query
    if ($archiveUrl) {
        return @{
            Url = $archiveUrl
            Source = 'archive.org'
        }
    }

    return $null
}

function Set-EntryPreviewFromImage(
    [pscustomobject]$entry,
    [string]$sourceImagePath,
    [string]$previewSource,
    [string]$previewImageUrl,
    [System.Windows.Forms.ListView]$listView,
    [System.Windows.Forms.ImageList]$tileImageList,
    [System.Windows.Forms.TextBoxBase]$detailsBox,
    [System.Windows.Forms.Label]$statsLabel,
    [System.Windows.Forms.TextBox]$logBox
) {
    if (-not $entry -or -not (Test-Path $entry.FolderPath)) {
        Show-Error 'Spielordner nicht gefunden.'
        return
    }

    $previewPath = Get-PreviewCachePath -folderPath $entry.FolderPath
    $ok = Convert-ImageToPreview -sourceImagePath $sourceImagePath -outputImagePath $previewPath
    if (-not $ok) {
        Show-Error 'Bild konnte nicht als Vorschau verarbeitet werden.'
        return
    }

    if ([string]::IsNullOrWhiteSpace($previewSource)) {
        $previewSource = 'manual'
    }

    $entry.PreviewImagePath = $previewPath
    $entry.PreviewSource = $previewSource
    $entry.PreviewImageUrl = if ($previewImageUrl) { $previewImageUrl } else { '' }
    $entry.PreviewUpdatedAt = (Get-Date).ToString('s')

    $script:PreviewsEnabled = $true
    $imgKey = Get-StringHash -text $entry.FolderPath
    if ($tileImageList -and $tileImageList.Images.ContainsKey($imgKey)) {
        $tileImageList.Images.RemoveByKey($imgKey)
    }

    if ($tileImageList) {
        [void](Ensure-EntryPreview -entry $entry -imageList $tileImageList -forceRebuild $false -logBox $null)
    }

    Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $tileImageList -logBox $logBox -entries $script:Entries
    $logBox.AppendText("Vorschau gesetzt: $($entry.Name) -> $previewPath`r`n")
}

function Get-GameExecutable([string]$folderPath) {
    $preferred = @(
        '!start.bat', '!run.bat', 'start.bat', 'run.bat', 'setup.bat',
        'install.exe', 'game.exe', 'play.exe', 'dos4gw.exe'
    )

    $allFiles = Get-ChildItem -Path $folderPath -File -ErrorAction SilentlyContinue

    foreach ($name in $preferred) {
        $match = $allFiles | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
        if ($match) {
            return $match
        }
    }

    $candidates = $allFiles |
        Where-Object { $_.Extension -match '^(\.bat|\.com|\.exe)$' } |
        Where-Object { $_.Name -notmatch '(?i)(unins|uninstall|setup|config|readme|dosbox)' } |
        Sort-Object Name

    return $candidates | Select-Object -First 1
}

function Get-OptimizationInfo([string]$folderPath) {
    $confFile = Get-ChildItem -Path $folderPath -File -Filter '*.conf' -ErrorAction SilentlyContinue |
        Sort-Object Name |
        Select-Object -First 1

    if ($confFile) {
        return @{
            IsOptimized = $true
            ConfigPath = $confFile.FullName
            Reason = "Konfig vorhanden: $($confFile.Name)"
            Source = 'existing'
        }
    }

    $launcherBat = Get-ChildItem -Path $folderPath -File -Filter '*.bat' -ErrorAction SilentlyContinue |
        Where-Object {
            try {
                (Get-Content -Path $_.FullName -Raw -ErrorAction Stop) -match '(?i)dosbox'
            } catch {
                $false
            }
        } |
        Select-Object -First 1

    if ($launcherBat) {
        return @{
            IsOptimized = $true
            ConfigPath = $null
            Reason = "DOSBox-Startskript: $($launcherBat.Name)"
            Source = 'existing'
        }
    }

    return @{
        IsOptimized = $false
        ConfigPath = $null
        Reason = 'Keine DOSBox-spezifische Konfiguration gefunden'
        Source = 'none'
    }
}

function Build-GameEntry([System.IO.DirectoryInfo]$dir, [string]$rootFolder) {
    $exe = Get-GameExecutable -folderPath $dir.FullName
    $opt = Get-OptimizationInfo -folderPath $dir.FullName
    $displayName = $dir.Name
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        $displayName = Split-Path -Leaf (($dir.FullName -replace '[\\/]+$',''))
    }
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        $displayName = '(Unbenanntes Spiel)'
    }

    return [PSCustomObject]@{
        Name = $displayName
        FolderPath = $dir.FullName
        RootFolder = $rootFolder
        RelativePath = $dir.Name
        ExecutableName = if ($exe) { $exe.Name } else { '' }
        ExecutableFullPath = if ($exe) { $exe.FullName } else { '' }
        IsOptimized = [bool]$opt.IsOptimized
        ConfigPath = $opt.ConfigPath
        OptimizationReason = $opt.Reason
        OptimizationSource = $opt.Source
        PreviewImagePath = ''
        PreviewSource = 'none'
        PreviewImageUrl = ''
        PreviewUpdatedAt = ''
        IsFavorite = $false
    }
}

function Build-RootGameEntry([string]$rootFolder) {
    $exe = Get-GameExecutable -folderPath $rootFolder
    $opt = Get-OptimizationInfo -folderPath $rootFolder
    $folderName = Split-Path -Leaf (($rootFolder -replace '[\\/]+$',''))
    if ([string]::IsNullOrWhiteSpace($folderName)) {
        $folderName = '(Root-Spiel)'
    }

    return [PSCustomObject]@{
        Name = $folderName
        FolderPath = $rootFolder
        RootFolder = $rootFolder
        RelativePath = '.'
        ExecutableName = if ($exe) { $exe.Name } else { '' }
        ExecutableFullPath = if ($exe) { $exe.FullName } else { '' }
        IsOptimized = [bool]$opt.IsOptimized
        ConfigPath = $opt.ConfigPath
        OptimizationReason = $opt.Reason
        OptimizationSource = $opt.Source
        PreviewImagePath = ''
        PreviewSource = 'none'
        PreviewImageUrl = ''
        PreviewUpdatedAt = ''
        IsFavorite = $false
    }
}

function Get-Games([string]$rootFolder) {
    $items = New-Object System.Collections.Generic.List[object]

    $rootEntry = Build-RootGameEntry -rootFolder $rootFolder
    if ($rootEntry.ExecutableName -or $rootEntry.IsOptimized) {
        [void]$items.Add($rootEntry)
    }

    $dirs = Get-ChildItem -Path $rootFolder -Directory -ErrorAction SilentlyContinue | Sort-Object Name
    foreach ($dir in $dirs) {
        $entry = Build-GameEntry -dir $dir -rootFolder $rootFolder
        [void]$items.Add($entry)
    }

    $rootExecutables = Get-ChildItem -Path $rootFolder -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -match '^(\.bat|\.com|\.exe)$' } |
        Where-Object { $_.Name -notmatch '(?i)(unins|uninstall|setup|config|readme|dosbox)' } |
        Sort-Object Name

    foreach ($exe in $rootExecutables) {
        if ($rootEntry.ExecutableName -and ($exe.Name -ieq $rootEntry.ExecutableName)) {
            continue
        }

        $opt = Get-OptimizationInfo -folderPath $rootFolder
        $entry = [PSCustomObject]@{
            Name = "(Root) $($exe.BaseName)"
            FolderPath = $rootFolder
            RootFolder = $rootFolder
            RelativePath = '.'
            ExecutableName = $exe.Name
            ExecutableFullPath = $exe.FullName
            IsOptimized = [bool]$opt.IsOptimized
            ConfigPath = $opt.ConfigPath
            OptimizationReason = $opt.Reason
            OptimizationSource = $opt.Source
            PreviewImagePath = ''
            PreviewSource = 'none'
            PreviewImageUrl = ''
            PreviewUpdatedAt = ''
            IsFavorite = $false
        }
        [void]$items.Add($entry)
    }

    return $items
}

function Get-RelativeDosFolder([string]$baseFolder, [string]$targetFolder) {
    $baseFull = [System.IO.Path]::GetFullPath($baseFolder).TrimEnd('\\')
    $targetFull = [System.IO.Path]::GetFullPath($targetFolder).TrimEnd('\\')

    if (-not $targetFull.StartsWith($baseFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }

    $relative = $targetFull.Substring($baseFull.Length).TrimStart('\\')
    return ($relative -replace '\\', '/')
}

function Optimize-GameEntry(
    [pscustomobject]$entry,
    [System.Windows.Forms.Form]$form,
    [System.Windows.Forms.TextBox]$logBox
) {
    if (-not (Test-Path $entry.FolderPath)) {
        Show-Error "Spielordner nicht gefunden: $($entry.FolderPath)"
        return $false
    }

    $startFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $startFolderDialog.Description = 'Waehle den Ordner, in dem die eigentliche Spiel-Startdatei liegt.'
    $startFolderDialog.ShowNewFolderButton = $false
    $startFolderDialog.SelectedPath = $entry.FolderPath

    if ($startFolderDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return $false
    }

    $selectedStartFolder = $startFolderDialog.SelectedPath

    $startFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $startFileDialog.Title = 'Waehle die echte Spiel-Startdatei'
    $startFileDialog.InitialDirectory = $selectedStartFolder
    $startFileDialog.Filter = 'Startdateien (*.bat;*.com;*.exe)|*.bat;*.com;*.exe|Alle Dateien (*.*)|*.*'
    $startFileDialog.Multiselect = $false

    $startFile = $null
    if ($startFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $picked = Get-Item -LiteralPath $startFileDialog.FileName -ErrorAction SilentlyContinue
        if ($picked -and $picked.PSIsContainer -eq $false) {
            $startFile = $picked
            $selectedStartFolder = Split-Path -Parent $picked.FullName
        }
    }

    if (-not $startFile) {
        $startFile = Get-GameExecutable -folderPath $selectedStartFolder
    }

    if (-not $startFile) {
        Show-Error 'Im gewaehlten Ordner wurde keine startbare Datei (.bat/.com/.exe) gefunden.'
        return $false
    }

    $relativeStartFolder = Get-RelativeDosFolder -baseFolder $entry.FolderPath -targetFolder $selectedStartFolder
    if ($null -eq $relativeStartFolder) {
        Show-Error 'Der ausgewaehlte Startordner muss innerhalb des Spielordners liegen.'
        return $false
    }

    $configPath = Join-Path $entry.FolderPath 'dosbox-game.conf'
    $startBatPath = Join-Path $entry.FolderPath '!start.bat'

    $mountRootEscaped = $selectedStartFolder.Replace('"', '""')
    $lines = @(
        '[sdl]',
        'fullscreen=false',
        'fulldouble=true',
        'fullresolution=desktop',
        'windowresolution=1280x960',
        'output=opengl',
        '',
        '[render]',
        'aspect=true',
        'scaler=normal2x',
        '',
        '[cpu]',
        'core=auto',
        'cycles=max',
        '',
        '[mixer]',
        'rate=44100',
        '',
        '[autoexec]',
        ('mount c "{0}"' -f $mountRootEscaped),
        'c:'
    )

    $lines += $startFile.Name

    Set-Content -Path $configPath -Value $lines -Encoding ASCII

    $batLines = @(
        '@echo off',
        'setlocal',
        ('"{0}" -conf "{1}"' -f $script:DosBoxExe, $configPath)
    )
    Set-Content -Path $startBatPath -Value $batLines -Encoding ASCII

    $entry.ExecutableName = '!start.bat'
    $entry.ExecutableFullPath = $startBatPath
    $entry.ConfigPath = $configPath
    $entry.IsOptimized = $true
    $entry.OptimizationReason = "Optimiert fuer Startdatei: $($startFile.Name) in $relativeStartFolder"
    $entry.OptimizationSource = 'ui'

    $logBox.AppendText("Optimiert: $($entry.Name) -> $configPath`r`n")
    $logBox.AppendText("Starter erstellt: $startBatPath`r`n")
    return $true
}

function Ensure-GameConfig([pscustomobject]$entry) {
    if ($entry.ConfigPath -and (Test-Path $entry.ConfigPath)) {
        return $entry.ConfigPath
    }

    $configPath = Join-Path $entry.FolderPath 'dosbox-game.conf'

    $mountRoot = $entry.RootFolder
    $mountRootEscaped = $mountRoot.Replace('"', '""')

    $relative = $entry.RelativePath
    $exeName = $entry.ExecutableName

    $lines = @(
        '[sdl]',
        'fullscreen=false',
        'fulldouble=true',
        'fullresolution=desktop',
        'windowresolution=1280x960',
        'output=opengl',
        '',
        '[render]',
        'aspect=true',
        'scaler=normal2x',
        '',
        '[cpu]',
        'core=auto',
        'cycles=max',
        '',
        '[mixer]',
        'rate=44100',
        '',
        '[autoexec]',
        ('mount c "{0}"' -f $mountRootEscaped)
    )

    if ($relative -and $relative -ne '.') {
        $lines += 'c:'
        $lines += ('cd "{0}"' -f $relative)
    } else {
        $lines += 'c:'
    }

    if ($exeName) {
        $lines += $exeName
    }

    Set-Content -Path $configPath -Value $lines -Encoding ASCII
    return $configPath
}

function Start-Game([pscustomobject]$entry, [System.Windows.Forms.TextBox]$logBox) {
    if (-not (Test-Path $script:DosBoxExe)) {
        Show-Error "DOSBox.exe wurde nicht gefunden: $script:DosBoxExe"
        return
    }

    if (-not $entry.ExecutableName) {
        Show-Error 'Fuer dieses Spiel wurde keine startbare Datei (.bat/.com/.exe) gefunden.'
        return
    }

    $directStartPath = $entry.ExecutableFullPath
    if ([string]::IsNullOrWhiteSpace($directStartPath) -and $entry.ExecutableName) {
        $candidate = Join-Path $entry.FolderPath $entry.ExecutableName
        if (Test-Path $candidate) {
            $directStartPath = $candidate
        }
    }

    if ($entry.IsOptimized -and $directStartPath -and (Test-Path $directStartPath)) {
        try {
            $ext = [System.IO.Path]::GetExtension($directStartPath)
            if ($ext -and ($ext.Equals('.bat', [System.StringComparison]::OrdinalIgnoreCase) -or $ext.Equals('.cmd', [System.StringComparison]::OrdinalIgnoreCase))) {
                Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', ('"{0}"' -f $directStartPath) -WorkingDirectory $entry.FolderPath | Out-Null
            } else {
                Start-Process -FilePath $directStartPath -WorkingDirectory $entry.FolderPath | Out-Null
            }
            $logBox.AppendText("Gestartet (direkt): $($entry.Name) -> $(Split-Path -Leaf $directStartPath)`r`n")
            return
        } catch {
            $logBox.AppendText("Direktstart fehlgeschlagen, versuche DOSBox-Aufruf: $($_.Exception.Message)`r`n")
        }
    }

    $dosboxArgs = @()

    if ($entry.ConfigPath -and (Test-Path $entry.ConfigPath)) {
        $dosboxArgs += '-conf'
        $dosboxArgs += $entry.ConfigPath
    } else {
        $dosboxArgs += '-c'
        $dosboxArgs += ('mount c "{0}"' -f $entry.RootFolder)
        $dosboxArgs += '-c'
        $dosboxArgs += 'c:'
        if ($entry.RelativePath -and $entry.RelativePath -ne '.') {
            $dosboxArgs += '-c'
            $dosboxArgs += ('cd "{0}"' -f $entry.RelativePath)
        }
        $dosboxArgs += '-c'
        $dosboxArgs += $entry.ExecutableName
    }

    try {
        Start-Process -FilePath $script:DosBoxExe -ArgumentList $dosboxArgs -WorkingDirectory $script:BaseDir | Out-Null
        $logBox.AppendText("Gestartet: $($entry.Name)`r`n")
    } catch {
        Show-Error "Start fehlgeschlagen: $($_.Exception.Message)"
    }
}

function Populate-List(
    [System.Windows.Forms.ListView]$listView,
    [System.Windows.Forms.Label]$statsLabel,
    [System.Windows.Forms.TextBoxBase]$detailsBox,
    [System.Windows.Forms.ImageList]$tileImages,
    [System.Windows.Forms.TextBox]$logBox,
    [object[]]$entries
) {
    $script:Entries = @($entries)

    # Filter und Sortierung aufbauen
    $displayEntries = @($script:Entries)

    if ($script:ShowFavoritesOnly) {
        $displayEntries = @($displayEntries | Where-Object { $_.IsFavorite -eq $true })
    }

    if (-not [string]::IsNullOrWhiteSpace($script:FilterText)) {
        $needle = $script:FilterText.ToLowerInvariant()
        $displayEntries = @($displayEntries | Where-Object { $_.Name -and $_.Name.ToLowerInvariant().Contains($needle) })
    }

    switch ($script:SortMode) {
        'name-asc'  { $displayEntries = @($displayEntries | Sort-Object Name) }
        'name-desc' { $displayEntries = @($displayEntries | Sort-Object Name -Descending) }
        'status'    { $displayEntries = @($displayEntries | Sort-Object { if ($_.IsOptimized) { 0 } else { 1 } }, Name) }
        'fav-first' { $displayEntries = @($displayEntries | Sort-Object { if ($_.IsFavorite) { 0 } else { 1 } }, Name) }
    }

    $listView.BeginUpdate()
    $listView.Items.Clear()

    foreach ($entry in $displayEntries) {
        $status = 'Nicht optimiert'
        if ($entry.IsOptimized) {
            if ($entry.OptimizationSource -eq 'ui') {
                $status = 'UI-Optimiert'
            } else {
                $status = 'Optimiert'
            }
        }
        $sub = if ($entry.ExecutableName) { $entry.ExecutableName } else { '(keine EXE/BAT/COM gefunden)' }
        $displayName = $entry.Name
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $displayName = Split-Path -Leaf (($entry.FolderPath -replace '[\\/]+$',''))
        }
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            if ($entry.ExecutableName) {
                $displayName = "Start: $($entry.ExecutableName)"
            } else {
                $displayName = '(Unbenanntes Spiel)'
            }
        }

        if ($entry.IsFavorite) {
            $displayName = "$([char]0x2605) $displayName"
        }

        $item = New-Object System.Windows.Forms.ListViewItem($displayName)
        [void]$item.SubItems.Add($sub)
        [void]$item.SubItems.Add($status)
        $item.Tag = $entry
        $item.ForeColor = [System.Drawing.Color]::Black
        if ($tileImages) {
            $item.ImageKey = 'default'
        }
        if ($script:PreviewsEnabled -and $tileImages) {
            try {
                $item.ImageKey = Ensure-EntryPreview -entry $entry -imageList $tileImages -forceRebuild $false -logBox $null
            } catch {
                if ($tileImages) {
                    $item.ImageKey = 'default'
                }
                if ($logBox) {
                    $logBox.AppendText("Preview-Fehler (ignoriert): $($entry.Name) - $($_.Exception.Message)`r`n")
                }
            }
        }

        if ($entry.IsFavorite) {
            $item.BackColor = [System.Drawing.Color]::FromArgb(255, 250, 205)
        } elseif ($entry.IsOptimized) {
            $item.BackColor = [System.Drawing.Color]::FromArgb(228, 249, 234)
        } else {
            $item.BackColor = [System.Drawing.Color]::FromArgb(255, 242, 230)
        }

        [void]$listView.Items.Add($item)
    }

    $listView.EndUpdate()

    $optimized = @($script:Entries | Where-Object { $_.IsOptimized -eq $true }).Count
    $favorites = @($script:Entries | Where-Object { $_.IsFavorite -eq $true }).Count
    $total = $script:Entries.Count
    $shown = $displayEntries.Count
    $filterInfo = if ($shown -lt $total) { " | Angezeigt: $shown" } else { '' }
    $statsLabel.Text = "Spiele: $total | Favoriten: $favorites | Optimiert: $optimized$filterInfo"
    $detailsBox.Text = ''

    if ($listView.Items.Count -gt 0) {
        $listView.Items[0].Selected = $true
        $listView.Select()
    }
}

function Get-SelectedEntry([System.Windows.Forms.ListView]$listView) {
    if ($listView.SelectedItems.Count -eq 0) {
        return $null
    }
    return $listView.SelectedItems[0].Tag
}

function Set-ScanUiState(
    [System.Windows.Forms.Form]$form,
    [System.Windows.Forms.Button]$browseButton,
    [System.Windows.Forms.Button]$rescanButton,
    [System.Windows.Forms.Button]$viewToggle,
    [System.Windows.Forms.Label]$statsLabel,
    [bool]$isBusy
) {
    $script:IsScanning = $isBusy
    $browseButton.Enabled = -not $isBusy
    $rescanButton.Enabled = -not $isBusy
    $viewToggle.Enabled = -not $isBusy
    $form.UseWaitCursor = $isBusy
    if ($isBusy) {
        $statsLabel.Text = 'Scan laeuft... bitte warten.'
    }
}

function Start-Scan(
    [System.Windows.Forms.Form]$form,
    [System.Windows.Forms.ListView]$listView,
    [System.Windows.Forms.Label]$statsLabel,
    [System.Windows.Forms.TextBoxBase]$detailsBox,
    [System.Windows.Forms.TextBox]$logBox,
    [System.Windows.Forms.Button]$browseButton,
    [System.Windows.Forms.Button]$rescanButton,
    [System.Windows.Forms.Button]$viewToggle,
    [string]$rootFolder,
    [string]$reason
) {
    if ($script:IsScanning) {
        return
    }

    if (-not (Test-Path $rootFolder)) {
        Show-Error "Ordner nicht gefunden: $rootFolder"
        return
    }

    Set-ScanUiState -form $form -browseButton $browseButton -rescanButton $rescanButton -viewToggle $viewToggle -statsLabel $statsLabel -isBusy $true
    $logBox.AppendText("Scan gestartet ($reason): $rootFolder`r`n")

    try {
        $dirs = @(Get-ChildItem -Path $rootFolder -Directory -ErrorAction SilentlyContinue | Sort-Object Name)
        $rootExecutables = @(Get-ChildItem -Path $rootFolder -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '^(\.bat|\.com|\.exe)$' } |
            Where-Object { $_.Name -notmatch '(?i)(unins|uninstall|setup|config|readme|dosbox)' } |
            Sort-Object Name)

        $results = New-Object System.Collections.Generic.List[object]
        $totalDirs = $dirs.Count

        $rootEntry = Build-RootGameEntry -rootFolder $rootFolder
        if ($rootEntry.ExecutableName -or $rootEntry.IsOptimized) {
            [void]$results.Add($rootEntry)
        }

        for ($i = 0; $i -lt $totalDirs; $i++) {
            $dir = $dirs[$i]
            if ($null -eq $dir) {
                continue
            }

            try {
                $entry = Build-GameEntry -dir $dir -rootFolder $rootFolder
                [void]$results.Add($entry)
            } catch {
                $logBox.AppendText("Warnung: Ordner uebersprungen ($($dir.FullName)) - $($_.Exception.Message)`r`n")
            }

            if (($i % 5) -eq 0) {
                $statsLabel.Text = "Scan laeuft... $($i + 1) / $totalDirs Ordner"
                [System.Windows.Forms.Application]::DoEvents()
            }
        }

        $rootOpt = Get-OptimizationInfo -folderPath $rootFolder
        foreach ($exe in $rootExecutables) {
            if ($rootEntry.ExecutableName -and ($exe.Name -ieq $rootEntry.ExecutableName)) {
                continue
            }

            $entry = [PSCustomObject]@{
                Name = "(Root) $($exe.BaseName)"
                FolderPath = $rootFolder
                RootFolder = $rootFolder
                RelativePath = '.'
                ExecutableName = $exe.Name
                ExecutableFullPath = $exe.FullName
                IsOptimized = [bool]$rootOpt.IsOptimized
                ConfigPath = $rootOpt.ConfigPath
                OptimizationReason = $rootOpt.Reason
                OptimizationSource = $rootOpt.Source
                PreviewImagePath = ''
                PreviewSource = 'none'
                PreviewImageUrl = ''
                PreviewUpdatedAt = ''
                IsFavorite = $false
            }
            [void]$results.Add($entry)
        }

        Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $tileImageList -logBox $logBox -entries $results.ToArray()
        [void](Save-FolderCache -rootFolder $rootFolder -entries $script:Entries)
        $logBox.AppendText("Scan abgeschlossen: $rootFolder`r`n")
        $logBox.AppendText("In Liste geladen: $($results.Count) Eintraege`r`n")
        
        # Lade Vorschaubilder nach Scan
        if ($results.Count -gt 0) {
            $ensureList = Ensure-TileImageList -listView $listView
            if ($ensureList) {
                $script:PreviewsEnabled = $true
                $logBox.AppendText("Lade Vorschaubilder nach Scan...`r`n")
                [System.Windows.Forms.Application]::DoEvents()
                
                $previewsLoaded = 0
                foreach ($entry in $script:Entries) {
                    try {
                        $key = Ensure-EntryPreview -entry $entry -imageList $ensureList -forceRebuild $false -logBox $null
                        if ($key -and $key -ne 'default') {
                            $previewsLoaded++
                        }
                    } catch {
                        # Silent continue
                    }
                    [System.Windows.Forms.Application]::DoEvents()
                }
                
                Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $ensureList -logBox $logBox -entries $script:Entries
                $logBox.AppendText("Vorschaubilder geladen: $previewsLoaded`r`n")
            }
        }
    } catch {
        Show-Error "Scan fehlgeschlagen: $($_.Exception.Message)"
        $logBox.AppendText("Scan fehlgeschlagen: $($_.Exception.Message)`r`n")
    } finally {
        Set-ScanUiState -form $form -browseButton $browseButton -rescanButton $rescanButton -viewToggle $viewToggle -statsLabel $statsLabel -isBusy $false
    }
}

if (-not (Test-Path $script:DosBoxExe)) {
    [System.Windows.Forms.MessageBox]::Show(
        "DOSBox.exe wurde im Ordner nicht gefunden.`r`nPfad: $script:DosBoxExe",
        'DOSBox Spiele-Manager',
        'OK',
        'Error'
    ) | Out-Null
    exit 1
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'DOSBox Spiele-Manager'
$form.Size = New-Object System.Drawing.Size(1380, 760)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(1260, 700)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 248)

$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = 'Top'
$topPanel.Height = 165
$topPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 43, 56)
$form.Controls.Add($topPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'DOSBox Spiele-Bibliothek'
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(18, 12)
$topPanel.Controls.Add($titleLabel)

$pathBox = New-Object System.Windows.Forms.TextBox
$pathBox.Location = New-Object System.Drawing.Point(18, 52)
$pathBox.Size = New-Object System.Drawing.Size(650, 26)
$pathBox.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$pathBox.ReadOnly = $false
$topPanel.Controls.Add($pathBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = 'Ordner waehlen'
$browseButton.Location = New-Object System.Drawing.Point(680, 49)
$browseButton.Size = New-Object System.Drawing.Size(130, 32)
$browseButton.FlatStyle = 'Flat'
$browseButton.BackColor = [System.Drawing.Color]::FromArgb(89, 139, 193)
$browseButton.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($browseButton)

$rescanButton = New-Object System.Windows.Forms.Button
$rescanButton.Text = 'Neu scannen'
$rescanButton.Location = New-Object System.Drawing.Point(820, 49)
$rescanButton.Size = New-Object System.Drawing.Size(110, 32)
$rescanButton.FlatStyle = 'Flat'
$rescanButton.BackColor = [System.Drawing.Color]::FromArgb(74, 116, 165)
$rescanButton.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($rescanButton)

$viewToggle = New-Object System.Windows.Forms.Button
$viewToggle.Text = 'Ansicht: Liste'
$viewToggle.Location = New-Object System.Drawing.Point(940, 49)
$viewToggle.Size = New-Object System.Drawing.Size(130, 32)
$viewToggle.FlatStyle = 'Flat'
$viewToggle.BackColor = [System.Drawing.Color]::FromArgb(60, 87, 123)
$viewToggle.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($viewToggle)

$saveCacheButton = New-Object System.Windows.Forms.Button
$saveCacheButton.Text = 'Daten speichern'
$saveCacheButton.Location = New-Object System.Drawing.Point(1080, 49)
$saveCacheButton.Size = New-Object System.Drawing.Size(125, 32)
$saveCacheButton.FlatStyle = 'Flat'
$saveCacheButton.BackColor = [System.Drawing.Color]::FromArgb(57, 138, 121)
$saveCacheButton.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($saveCacheButton)

$loadCacheButton = New-Object System.Windows.Forms.Button
$loadCacheButton.Text = 'Daten laden'
$loadCacheButton.Location = New-Object System.Drawing.Point(1210, 49)
$loadCacheButton.Size = New-Object System.Drawing.Size(120, 32)
$loadCacheButton.FlatStyle = 'Flat'
$loadCacheButton.BackColor = [System.Drawing.Color]::FromArgb(77, 120, 153)
$loadCacheButton.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($loadCacheButton)

$buildPreviewsButton = New-Object System.Windows.Forms.Button
$buildPreviewsButton.Text = 'Vorschauen erstellen'
$buildPreviewsButton.Location = New-Object System.Drawing.Point(1080, 12)
$buildPreviewsButton.Size = New-Object System.Drawing.Size(250, 30)
$buildPreviewsButton.FlatStyle = 'Flat'
$buildPreviewsButton.BackColor = [System.Drawing.Color]::FromArgb(122, 88, 160)
$buildPreviewsButton.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($buildPreviewsButton)

$statsLabel = New-Object System.Windows.Forms.Label
$statsLabel.Text = 'Spiele: 0'
$statsLabel.AutoSize = $true
$statsLabel.Location = New-Object System.Drawing.Point(20, 94)
$statsLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$statsLabel.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($statsLabel)

$filterLabel = New-Object System.Windows.Forms.Label
$filterLabel.Text = 'Suche:'
$filterLabel.AutoSize = $true
$filterLabel.Location = New-Object System.Drawing.Point(20, 134)
$filterLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$filterLabel.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($filterLabel)

$filterBox = New-Object System.Windows.Forms.TextBox
$filterBox.Location = New-Object System.Drawing.Point(68, 131)
$filterBox.Size = New-Object System.Drawing.Size(270, 26)
$filterBox.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$topPanel.Controls.Add($filterBox)

$clearFilterButton = New-Object System.Windows.Forms.Button
$clearFilterButton.Text = 'X'
$clearFilterButton.Location = New-Object System.Drawing.Point(344, 131)
$clearFilterButton.Size = New-Object System.Drawing.Size(26, 26)
$clearFilterButton.FlatStyle = 'Flat'
$clearFilterButton.BackColor = [System.Drawing.Color]::FromArgb(89, 139, 193)
$clearFilterButton.ForeColor = [System.Drawing.Color]::White
$clearFilterButton.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$topPanel.Controls.Add($clearFilterButton)

$favFilterButton = New-Object System.Windows.Forms.Button
$favFilterButton.Text = "$([char]0x2606) Nur Favoriten"
$favFilterButton.Location = New-Object System.Drawing.Point(378, 131)
$favFilterButton.Size = New-Object System.Drawing.Size(148, 26)
$favFilterButton.FlatStyle = 'Flat'
$favFilterButton.BackColor = [System.Drawing.Color]::FromArgb(140, 110, 40)
$favFilterButton.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($favFilterButton)

$sortButton = New-Object System.Windows.Forms.Button
$sortButton.Text = 'Sortierung: A-Z'
$sortButton.Location = New-Object System.Drawing.Point(534, 131)
$sortButton.Size = New-Object System.Drawing.Size(150, 26)
$sortButton.FlatStyle = 'Flat'
$sortButton.BackColor = [System.Drawing.Color]::FromArgb(74, 116, 165)
$sortButton.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($sortButton)

$headerToggleButton = New-Object System.Windows.Forms.Button
$headerToggleButton.Text = 'Header ausblenden'
$headerToggleButton.Size = New-Object System.Drawing.Size(150, 28)
$headerToggleButton.Anchor = 'Top, Right'
$headerToggleButton.FlatStyle = 'Flat'
$headerToggleButton.BackColor = [System.Drawing.Color]::FromArgb(74, 116, 165)
$headerToggleButton.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($headerToggleButton)

$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$contentPanel.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
$contentPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 248)
$form.Controls.Add($contentPanel)

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'
$split.SplitterDistance = 720
$split.SplitterWidth = 8
$split.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 248)
$contentPanel.Controls.Add($split)

$listView = New-Object System.Windows.Forms.ListView
$listView.Dock = 'Fill'
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.HideSelection = $false
$listView.MultiSelect = $false
$listView.GridLines = $true
$listView.BackColor = [System.Drawing.Color]::White
$listView.ForeColor = [System.Drawing.Color]::Black
$listView.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$script:tileImageList = $null
[void]$listView.Columns.Add('Spiel', 250)
[void]$listView.Columns.Add('Startdatei', 220)
[void]$listView.Columns.Add('Status', 120)
$listView.TileSize = New-Object System.Drawing.Size(360, 130)
$split.Panel1.Controls.Add($listView)

$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Dock = 'Fill'
$rightPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 248)
$split.Panel2.Controls.Add($rightPanel)

$detailsLabel = New-Object System.Windows.Forms.Label
$detailsLabel.Text = 'Details'
$detailsLabel.AutoSize = $true
$detailsLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11, [System.Drawing.FontStyle]::Bold)
$detailsLabel.Location = New-Object System.Drawing.Point(10, 12)
$rightPanel.Controls.Add($detailsLabel)

$quickNameLabel = New-Object System.Windows.Forms.Label
$quickNameLabel.Text = 'Name: -'
$quickNameLabel.AutoEllipsis = $true
$quickNameLabel.Location = New-Object System.Drawing.Point(12, 36)
$quickNameLabel.Size = New-Object System.Drawing.Size(330, 18)
$quickNameLabel.Anchor = 'Top, Left, Right'
$quickNameLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$rightPanel.Controls.Add($quickNameLabel)

$quickExeLabel = New-Object System.Windows.Forms.Label
$quickExeLabel.Text = 'Startdatei: -'
$quickExeLabel.AutoEllipsis = $true
$quickExeLabel.Location = New-Object System.Drawing.Point(12, 54)
$quickExeLabel.Size = New-Object System.Drawing.Size(330, 18)
$quickExeLabel.Anchor = 'Top, Left, Right'
$quickExeLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$rightPanel.Controls.Add($quickExeLabel)

$quickStatusLabel = New-Object System.Windows.Forms.Label
$quickStatusLabel.Text = 'Optimiert: -'
$quickStatusLabel.AutoEllipsis = $true
$quickStatusLabel.Location = New-Object System.Drawing.Point(12, 72)
$quickStatusLabel.Size = New-Object System.Drawing.Size(330, 18)
$quickStatusLabel.Anchor = 'Top, Left, Right'
$quickStatusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$rightPanel.Controls.Add($quickStatusLabel)

$detailsBox = New-Object System.Windows.Forms.RichTextBox
$detailsBox.ReadOnly = $true
$detailsBox.WordWrap = $false
$detailsBox.ScrollBars = 'Vertical'
$detailsBox.DetectUrls = $false
$detailsBox.ShortcutsEnabled = $true
$detailsBox.Location = New-Object System.Drawing.Point(12, 96)
$detailsBox.Size = New-Object System.Drawing.Size(330, 202)
$detailsBox.Anchor = 'Top, Bottom, Left, Right'
$detailsBox.Font = New-Object System.Drawing.Font('Consolas', 10)
$detailsBox.BackColor = [System.Drawing.Color]::White
$detailsBox.ForeColor = [System.Drawing.Color]::Black
$detailsBox.BorderStyle = 'FixedSingle'
$detailsBox.Visible = $false
$rightPanel.Controls.Add($detailsBox)

$fieldNameLabel = New-Object System.Windows.Forms.Label
$fieldNameLabel.Text = 'Name:'
$fieldNameLabel.Location = New-Object System.Drawing.Point(12, 102)
$fieldNameLabel.Size = New-Object System.Drawing.Size(85, 20)
$fieldNameLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$rightPanel.Controls.Add($fieldNameLabel)

$fieldNameValue = New-Object System.Windows.Forms.TextBox
$fieldNameValue.Location = New-Object System.Drawing.Point(100, 100)
$fieldNameValue.Size = New-Object System.Drawing.Size(242, 22)
$fieldNameValue.Anchor = 'Top, Left, Right'
$fieldNameValue.ReadOnly = $true
$fieldNameValue.BackColor = [System.Drawing.Color]::White
$fieldNameValue.ForeColor = [System.Drawing.Color]::Black
$rightPanel.Controls.Add($fieldNameValue)

$fieldFolderLabel = New-Object System.Windows.Forms.Label
$fieldFolderLabel.Text = 'Ordner:'
$fieldFolderLabel.Location = New-Object System.Drawing.Point(12, 132)
$fieldFolderLabel.Size = New-Object System.Drawing.Size(85, 20)
$fieldFolderLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$rightPanel.Controls.Add($fieldFolderLabel)

$fieldFolderValue = New-Object System.Windows.Forms.TextBox
$fieldFolderValue.Location = New-Object System.Drawing.Point(100, 130)
$fieldFolderValue.Size = New-Object System.Drawing.Size(242, 22)
$fieldFolderValue.Anchor = 'Top, Left, Right'
$fieldFolderValue.ReadOnly = $true
$fieldFolderValue.BackColor = [System.Drawing.Color]::White
$fieldFolderValue.ForeColor = [System.Drawing.Color]::Black
$rightPanel.Controls.Add($fieldFolderValue)

$fieldStartLabel = New-Object System.Windows.Forms.Label
$fieldStartLabel.Text = 'Startdatei:'
$fieldStartLabel.Location = New-Object System.Drawing.Point(12, 162)
$fieldStartLabel.Size = New-Object System.Drawing.Size(85, 20)
$fieldStartLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$rightPanel.Controls.Add($fieldStartLabel)

$fieldStartValue = New-Object System.Windows.Forms.TextBox
$fieldStartValue.Location = New-Object System.Drawing.Point(100, 160)
$fieldStartValue.Size = New-Object System.Drawing.Size(242, 22)
$fieldStartValue.Anchor = 'Top, Left, Right'
$fieldStartValue.ReadOnly = $true
$fieldStartValue.BackColor = [System.Drawing.Color]::White
$fieldStartValue.ForeColor = [System.Drawing.Color]::Black
$rightPanel.Controls.Add($fieldStartValue)

$fieldOptimizedLabel = New-Object System.Windows.Forms.Label
$fieldOptimizedLabel.Text = 'Optimiert:'
$fieldOptimizedLabel.Location = New-Object System.Drawing.Point(12, 192)
$fieldOptimizedLabel.Size = New-Object System.Drawing.Size(85, 20)
$fieldOptimizedLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$rightPanel.Controls.Add($fieldOptimizedLabel)

$fieldOptimizedValue = New-Object System.Windows.Forms.TextBox
$fieldOptimizedValue.Location = New-Object System.Drawing.Point(100, 190)
$fieldOptimizedValue.Size = New-Object System.Drawing.Size(242, 22)
$fieldOptimizedValue.Anchor = 'Top, Left, Right'
$fieldOptimizedValue.ReadOnly = $true
$fieldOptimizedValue.BackColor = [System.Drawing.Color]::White
$fieldOptimizedValue.ForeColor = [System.Drawing.Color]::Black
$rightPanel.Controls.Add($fieldOptimizedValue)

$fieldReasonLabel = New-Object System.Windows.Forms.Label
$fieldReasonLabel.Text = 'Erkennung:'
$fieldReasonLabel.Location = New-Object System.Drawing.Point(12, 222)
$fieldReasonLabel.Size = New-Object System.Drawing.Size(85, 20)
$fieldReasonLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$rightPanel.Controls.Add($fieldReasonLabel)

$fieldReasonValue = New-Object System.Windows.Forms.TextBox
$fieldReasonValue.Location = New-Object System.Drawing.Point(100, 220)
$fieldReasonValue.Size = New-Object System.Drawing.Size(242, 22)
$fieldReasonValue.Anchor = 'Top, Left, Right'
$fieldReasonValue.ReadOnly = $true
$fieldReasonValue.BackColor = [System.Drawing.Color]::White
$fieldReasonValue.ForeColor = [System.Drawing.Color]::Black
$rightPanel.Controls.Add($fieldReasonValue)

$fieldConfigLabel = New-Object System.Windows.Forms.Label
$fieldConfigLabel.Text = 'Config:'
$fieldConfigLabel.Location = New-Object System.Drawing.Point(12, 252)
$fieldConfigLabel.Size = New-Object System.Drawing.Size(85, 20)
$fieldConfigLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$rightPanel.Controls.Add($fieldConfigLabel)

$fieldConfigValue = New-Object System.Windows.Forms.TextBox
$fieldConfigValue.Location = New-Object System.Drawing.Point(100, 250)
$fieldConfigValue.Size = New-Object System.Drawing.Size(242, 22)
$fieldConfigValue.Anchor = 'Top, Left, Right'
$fieldConfigValue.ReadOnly = $true
$fieldConfigValue.BackColor = [System.Drawing.Color]::White
$fieldConfigValue.ForeColor = [System.Drawing.Color]::Black
$rightPanel.Controls.Add($fieldConfigValue)

$favToggleButton = New-Object System.Windows.Forms.Button
$favToggleButton.Text = "$([char]0x2606) Als Favorit markieren"
$favToggleButton.Location = New-Object System.Drawing.Point(12, 282)
$favToggleButton.Size = New-Object System.Drawing.Size(330, 26)
$favToggleButton.Anchor = 'Top, Left, Right'
$favToggleButton.FlatStyle = 'Flat'
$favToggleButton.BackColor = [System.Drawing.Color]::FromArgb(140, 110, 40)
$favToggleButton.ForeColor = [System.Drawing.Color]::White
$rightPanel.Controls.Add($favToggleButton)

$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = 'Spiel starten'
$startButton.Location = New-Object System.Drawing.Point(12, 310)
$startButton.Size = New-Object System.Drawing.Size(160, 36)
$startButton.Anchor = 'Bottom, Left'
$startButton.FlatStyle = 'Flat'
$startButton.BackColor = [System.Drawing.Color]::FromArgb(58, 153, 90)
$startButton.ForeColor = [System.Drawing.Color]::White
$rightPanel.Controls.Add($startButton)

$optimizeButton = New-Object System.Windows.Forms.Button
$optimizeButton.Text = 'Optimieren'
$optimizeButton.Location = New-Object System.Drawing.Point(182, 310)
$optimizeButton.Size = New-Object System.Drawing.Size(160, 36)
$optimizeButton.Anchor = 'Bottom, Right'
$optimizeButton.FlatStyle = 'Flat'
$optimizeButton.BackColor = [System.Drawing.Color]::FromArgb(211, 132, 38)
$optimizeButton.ForeColor = [System.Drawing.Color]::White
$rightPanel.Controls.Add($optimizeButton)

$editButton = New-Object System.Windows.Forms.Button
$editButton.Text = 'Anpassen (Config)'
$editButton.Location = New-Object System.Drawing.Point(12, 354)
$editButton.Size = New-Object System.Drawing.Size(160, 36)
$editButton.Anchor = 'Bottom, Left'
$editButton.FlatStyle = 'Flat'
$editButton.BackColor = [System.Drawing.Color]::FromArgb(74, 116, 165)
$editButton.ForeColor = [System.Drawing.Color]::White
$rightPanel.Controls.Add($editButton)

$optionsButton = New-Object System.Windows.Forms.Button
$optionsButton.Text = 'DOSBox Optionen'
$optionsButton.Location = New-Object System.Drawing.Point(182, 354)
$optionsButton.Size = New-Object System.Drawing.Size(160, 36)
$optionsButton.Anchor = 'Bottom, Right'
$optionsButton.FlatStyle = 'Flat'
$optionsButton.BackColor = [System.Drawing.Color]::FromArgb(109, 86, 165)
$optionsButton.ForeColor = [System.Drawing.Color]::White
$rightPanel.Controls.Add($optionsButton)

$addImageButton = New-Object System.Windows.Forms.Button
$addImageButton.Text = ('Bilder hinzuf' + $script:umlautu + 'gen')
$addImageButton.Location = New-Object System.Drawing.Point(12, 398)
$addImageButton.Size = New-Object System.Drawing.Size(330, 28)
$addImageButton.Anchor = 'Bottom, Left, Right'
$addImageButton.FlatStyle = 'Flat'
$addImageButton.BackColor = [System.Drawing.Color]::FromArgb(132, 95, 52)
$addImageButton.ForeColor = [System.Drawing.Color]::White
$rightPanel.Controls.Add($addImageButton)

$googleImageButton = New-Object System.Windows.Forms.Button
$googleImageButton.Text = 'Google-Bild'
$googleImageButton.Location = New-Object System.Drawing.Point(12, 430)
$googleImageButton.Size = New-Object System.Drawing.Size(330, 26)
$googleImageButton.Anchor = 'Bottom, Left, Right'
$googleImageButton.FlatStyle = 'Flat'
$googleImageButton.BackColor = [System.Drawing.Color]::FromArgb(96, 120, 58)
$googleImageButton.ForeColor = [System.Drawing.Color]::White
$rightPanel.Controls.Add($googleImageButton)

$googleImageBatchButton = New-Object System.Windows.Forms.Button
$googleImageBatchButton.Text = 'Google-Bilder (ALLE)'
$googleImageBatchButton.Location = New-Object System.Drawing.Point(12, 458)
$googleImageBatchButton.Size = New-Object System.Drawing.Size(330, 26)
$googleImageBatchButton.Anchor = 'Bottom, Left, Right'
$googleImageBatchButton.FlatStyle = 'Flat'
$googleImageBatchButton.BackColor = [System.Drawing.Color]::FromArgb(165, 97, 40)
$googleImageBatchButton.ForeColor = [System.Drawing.Color]::White
$rightPanel.Controls.Add($googleImageBatchButton)

$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Text = 'Log'
$logLabel.AutoSize = $true
$logLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10, [System.Drawing.FontStyle]::Bold)
$logLabel.Location = New-Object System.Drawing.Point(10, 488)
$rightPanel.Controls.Add($logLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ReadOnly = $true
$logBox.ScrollBars = 'Vertical'
$logBox.Location = New-Object System.Drawing.Point(12, 512)
$logBox.Size = New-Object System.Drawing.Size(330, 106)
$logBox.Anchor = 'Bottom, Left, Right'
$logBox.Font = New-Object System.Drawing.Font('Consolas', 8)
$logBox.BackColor = [System.Drawing.Color]::White
$logBox.ForeColor = [System.Drawing.Color]::Black
$logBox.BorderStyle = 'FixedSingle'
$rightPanel.Controls.Add($logBox)

$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$folderDialog.Description = 'Waehle den Hauptordner, in dem sich deine Spiele befinden.'
$folderDialog.ShowNewFolderButton = $false

$script:IsTileView = $false

$viewToggle.Add_Click({
    $script:IsTileView = -not $script:IsTileView
    if ($script:IsTileView) {
        $listView.View = 'Tile'
        $viewToggle.Text = 'Ansicht: Kacheln'
    } else {
        $listView.View = 'Details'
        $viewToggle.Text = 'Ansicht: Liste'
    }
})

$filterBox.Add_TextChanged({
    $script:FilterText = $filterBox.Text
    if ($script:Entries -and $script:Entries.Count -gt 0) {
        Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $script:tileImageList -logBox $logBox -entries $script:Entries
    }
})

$clearFilterButton.Add_Click({
    $filterBox.Text = ''
})

$favFilterButton.Add_Click({
    $script:ShowFavoritesOnly = -not $script:ShowFavoritesOnly
    if ($script:ShowFavoritesOnly) {
        $favFilterButton.BackColor = [System.Drawing.Color]::FromArgb(200, 160, 40)
        $favFilterButton.Text = "$([char]0x2605) Alle zeigen"
    } else {
        $favFilterButton.BackColor = [System.Drawing.Color]::FromArgb(140, 110, 40)
        $favFilterButton.Text = "$([char]0x2606) Nur Favoriten"
    }
    if ($script:Entries -and $script:Entries.Count -gt 0) {
        Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $script:tileImageList -logBox $logBox -entries $script:Entries
    }
})

$sortButton.Add_Click({
    switch ($script:SortMode) {
        'name-asc'  { $script:SortMode = 'name-desc'; $sortButton.Text = 'Sortierung: Z-A' }
        'name-desc' { $script:SortMode = 'status';    $sortButton.Text = 'Sortierung: Status' }
        'status'    { $script:SortMode = 'fav-first'; $sortButton.Text = 'Sortierung: Favoriten' }
        'fav-first' { $script:SortMode = 'name-asc';  $sortButton.Text = 'Sortierung: A-Z' }
    }
    if ($script:Entries -and $script:Entries.Count -gt 0) {
        Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $script:tileImageList -logBox $logBox -entries $script:Entries
    }
})

$saveCacheButton.Add_Click({
    if (-not $script:CurrentRoot) {
        Show-Error 'Bitte zuerst einen Spieleordner auswaehlen.'
        return
    }

    if (Save-FolderCache -rootFolder $script:CurrentRoot -entries $script:Entries) {
        $logBox.AppendText("Daten gespeichert: $(Get-CachePath -rootFolder $script:CurrentRoot)`r`n")
    } else {
        Show-Error 'Daten konnten nicht gespeichert werden.'
    }
})

$loadCacheButton.Add_Click({
    if (-not $script:CurrentRoot) {
        Show-Error 'Bitte zuerst einen Spieleordner auswaehlen.'
        return
    }

    $loaded = Load-FolderCache -rootFolder $script:CurrentRoot
    if ($null -eq $loaded) {
        Show-Error "Keine gueltigen gespeicherten Daten gefunden in: $(Get-CachePath -rootFolder $script:CurrentRoot)"
        return
    }

    # Initialisiere ImageList und lade Vorschaubilder
    $ensureList = Ensure-TileImageList -listView $listView
    if ($ensureList -and $loaded.Count -gt 0) {
        $script:PreviewsEnabled = $true
        $logBox.AppendText("Lade Vorschaubilder fuer Cache-Daten...`r`n")
        [System.Windows.Forms.Application]::DoEvents()
        
        $previewsLoaded = 0
        foreach ($entry in $loaded) {
            try {
                $key = Ensure-EntryPreview -entry $entry -imageList $ensureList -forceRebuild $false -logBox $null
                if ($key -and $key -ne 'default') {
                    $previewsLoaded++
                }
            } catch {
                # Silent continue
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        $logBox.AppendText("Vorschaubilder geladen: $previewsLoaded`r`n")
    }

    Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $ensureList -logBox $logBox -entries $loaded
    $logBox.AppendText("Daten geladen: $(Get-CachePath -rootFolder $script:CurrentRoot)`r`n")
})

$buildPreviewsButton.Add_Click({
    if (-not $script:CurrentRoot) {
        Show-Error 'Bitte zuerst einen Spieleordner auswaehlen.'
        return
    }

    if (-not $script:Entries -or $script:Entries.Count -eq 0) {
        Show-Error 'Keine Spiele geladen. Bitte zuerst scannen.'
        return
    }

    $tileImageList = Ensure-TileImageList -listView $listView
    if (-not $tileImageList) {
        Show-Error 'Preview-Bildliste konnte nicht initialisiert werden.'
        return
    }

    $script:PreviewsEnabled = $true
    $built = 0
    foreach ($entry in $script:Entries) {
        $key = Ensure-EntryPreview -entry $entry -imageList $tileImageList -forceRebuild $true -logBox $logBox
        if ($key -and $key -ne 'default') {
            $built++
        }
    }

    Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $tileImageList -logBox $logBox -entries $script:Entries
    $logBox.AppendText("Vorschauen erstellt/aktualisiert: $built`r`n")
})

$headerToggleButton.Add_Click({
    $topPanel.Visible = -not $topPanel.Visible
    if ($topPanel.Visible) {
        $headerToggleButton.Text = 'Header ausblenden'
    } else {
        $headerToggleButton.Text = 'Header einblenden'
    }
})

$form.Add_Resize({
    $headerToggleButton.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 165), 8)
    $headerToggleButton.BringToFront()

    $rightX = [Math]::Max(12, ($rightPanel.ClientSize.Width - 172))
    $optimizeButton.Location = New-Object System.Drawing.Point($rightX, 310)
    $optionsButton.Location = New-Object System.Drawing.Point($rightX, 354)
})

$pathBox.Add_KeyDown({
    param($source, $e)

    if ($e.KeyCode -ne [System.Windows.Forms.Keys]::Enter) {
        return
    }

    if ($script:IsScanning) {
        return
    }

    $candidate = $pathBox.Text.Trim()
    if (-not $candidate) {
        Show-Error 'Bitte einen gueltigen Ordnerpfad eingeben.'
        return
    }

    if (-not (Test-Path $candidate)) {
        Show-Error "Ordner nicht gefunden: $candidate"
        return
    }

    $script:CurrentRoot = $candidate
    Save-Settings -rootFolder $script:CurrentRoot
    Start-Scan -form $form -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -logBox $logBox -browseButton $browseButton -rescanButton $rescanButton -viewToggle $viewToggle -rootFolder $script:CurrentRoot -reason 'Pfad aus Textfeld'
    $logBox.AppendText("Ordner gesetzt (Textfeld): $($script:CurrentRoot)`r`n")
    $e.SuppressKeyPress = $true
})

$browseButton.Add_Click({
    if ($script:IsScanning) {
        return
    }

    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:CurrentRoot = $folderDialog.SelectedPath
        $pathBox.Text = $script:CurrentRoot
        Save-Settings -rootFolder $script:CurrentRoot
        Start-Scan -form $form -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -logBox $logBox -browseButton $browseButton -rescanButton $rescanButton -viewToggle $viewToggle -rootFolder $script:CurrentRoot -reason 'Ordnerauswahl'
        $logBox.AppendText("Ordner gesetzt: $($script:CurrentRoot)`r`n")
    }
})

$rescanButton.Add_Click({
    if ($script:IsScanning) {
        return
    }

    if (-not $script:CurrentRoot) {
        Show-Error 'Bitte zuerst einen Spieleordner auswaehlen.'
        return
    }
    Start-Scan -form $form -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -logBox $logBox -browseButton $browseButton -rescanButton $rescanButton -viewToggle $viewToggle -rootFolder $script:CurrentRoot -reason 'Manueller Rescan'
})

$listView.Add_SelectedIndexChanged({
    $entry = Get-SelectedEntry -listView $listView
    if (-not $entry) {
        $detailsBox.Text = ''
        $quickNameLabel.Text = 'Name: -'
        $quickExeLabel.Text = 'Startdatei: -'
        $quickStatusLabel.Text = 'Optimiert: -'
        $fieldNameValue.Text = ''
        $fieldFolderValue.Text = ''
        $fieldStartValue.Text = ''
        $fieldOptimizedValue.Text = ''
        $fieldReasonValue.Text = ''
        $fieldConfigValue.Text = ''
        $favToggleButton.Text = "$([char]0x2606) Als Favorit markieren"
        $favToggleButton.BackColor = [System.Drawing.Color]::FromArgb(140, 110, 40)
        return
    }

    $status = if ($entry.IsOptimized) { 'Ja' } else { 'Nein' }
    $statusMode = if ($entry.IsOptimized -and $entry.OptimizationSource -eq 'ui') { 'UI-Optimiert' } elseif ($entry.IsOptimized) { 'Optimiert' } else { 'Nicht optimiert' }
    $exe = if ($entry.ExecutableName) { $entry.ExecutableName } else { '(nicht gefunden)' }
    $conf = if ($entry.ConfigPath) { $entry.ConfigPath } else { '(keine)' }
    $nameValue = $entry.Name
    if ([string]::IsNullOrWhiteSpace($nameValue)) {
        $nameValue = Split-Path -Leaf (($entry.FolderPath -replace '[\\/]+$',''))
    }
    if ([string]::IsNullOrWhiteSpace($nameValue)) {
        if ($entry.ExecutableName) {
            $nameValue = "Start: $($entry.ExecutableName)"
        } else {
            $nameValue = '(Unbenanntes Spiel)'
        }
    }

    $quickNameLabel.Text = "Name: $nameValue"
    $quickExeLabel.Text = "Startdatei: $exe"
    $quickStatusLabel.Text = "Optimiert: $statusMode"

    $fieldNameValue.Text = $nameValue
    $fieldFolderValue.Text = $entry.FolderPath
    $fieldStartValue.Text = $exe
    $fieldOptimizedValue.Text = $statusMode
    $fieldReasonValue.Text = $entry.OptimizationReason
    $fieldConfigValue.Text = $conf

    $detailLines = @(
        "Name: $($entry.Name)",
        "Ordner: $($entry.FolderPath)",
        "Startdatei: $exe",
        "Optimiert: $statusMode",
        "Erkennung: $($entry.OptimizationReason)",
        "Config: $conf"
    )

    if ($entry.IsFavorite) {
        $favToggleButton.Text = "$([char]0x2605) Favorit entfernen"
        $favToggleButton.BackColor = [System.Drawing.Color]::FromArgb(200, 160, 40)
    } else {
        $favToggleButton.Text = "$([char]0x2606) Als Favorit markieren"
        $favToggleButton.BackColor = [System.Drawing.Color]::FromArgb(140, 110, 40)
    }

    $detailsBox.Lines = $detailLines

    # Always reset caret/scroll to top so details begin with Name/Ordner.
    $detailsBox.SelectionStart = 0
    $detailsBox.SelectionLength = 0
    $detailsBox.ScrollToCaret()
})

$startButton.Add_Click({
    $entry = Get-SelectedEntry -listView $listView
    if (-not $entry) {
        Show-Error 'Bitte zuerst ein Spiel waehlen.'
        return
    }
    Start-Game -entry $entry -logBox $logBox
})

$optimizeButton.Add_Click({
    $entry = Get-SelectedEntry -listView $listView
    if (-not $entry) {
        Show-Error 'Bitte zuerst ein Spiel waehlen.'
        return
    }

    try {
        if (Optimize-GameEntry -entry $entry -form $form -logBox $logBox) {
            Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $tileImageList -logBox $logBox -entries $script:Entries
            if ($script:CurrentRoot) {
                [void](Save-FolderCache -rootFolder $script:CurrentRoot -entries $script:Entries)
            }
        }
    } catch {
        Show-Error "Optimierung fehlgeschlagen: $($_.Exception.Message)"
    }
})

$editButton.Add_Click({
    $entry = Get-SelectedEntry -listView $listView
    if (-not $entry) {
        Show-Error 'Bitte zuerst ein Spiel waehlen.'
        return
    }

    try {
        $configPath = Ensure-GameConfig -entry $entry
        Start-Process -FilePath 'notepad.exe' -ArgumentList $configPath | Out-Null
        $logBox.AppendText("Config geoeffnet: $configPath`r`n")
    } catch {
        Show-Error "Config konnte nicht geoeffnet werden: $($_.Exception.Message)"
    }
})

$optionsButton.Add_Click({
    if (Test-Path $script:OptionsBat) {
        Start-Process -FilePath $script:OptionsBat -WorkingDirectory $script:BaseDir | Out-Null
    } else {
        Show-Error "Optionen-Datei nicht gefunden: $script:OptionsBat"
    }
})

$favToggleButton.Add_Click({
    $entry = Get-SelectedEntry -listView $listView
    if (-not $entry) {
        Show-Error 'Bitte zuerst ein Spiel waehlen.'
        return
    }
    $entry.IsFavorite = -not [bool]$entry.IsFavorite
    if ($entry.IsFavorite) {
        $favToggleButton.Text = "$([char]0x2605) Favorit entfernen"
        $favToggleButton.BackColor = [System.Drawing.Color]::FromArgb(200, 160, 40)
    } else {
        $favToggleButton.Text = "$([char]0x2606) Als Favorit markieren"
        $favToggleButton.BackColor = [System.Drawing.Color]::FromArgb(140, 110, 40)
    }
    if ($script:CurrentRoot) {
        [void](Save-FolderCache -rootFolder $script:CurrentRoot -entries $script:Entries)
    }
    Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $script:tileImageList -logBox $logBox -entries $script:Entries
})

$addImageButton.Add_Click({
    $entry = Get-SelectedEntry -listView $listView
    if (-not $entry) {
        Show-Error 'Bitte zuerst ein Spiel waehlen.'
        return
    }

    if (-not (Test-Path $entry.FolderPath)) {
        Show-Error "Spielordner nicht gefunden: $($entry.FolderPath)"
        return
    }

    $imageDialog = New-Object System.Windows.Forms.OpenFileDialog
    $imageDialog.Title = 'Vorschaubild waehlen'
    $imageDialog.InitialDirectory = $entry.FolderPath
    $imageDialog.Filter = 'Bilddateien (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif|Alle Dateien (*.*)|*.*'
    $imageDialog.Multiselect = $false

    if ($imageDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $tileImageList = Ensure-TileImageList -listView $listView
    if (-not $tileImageList) {
        Show-Error 'Bildliste konnte nicht initialisiert werden.'
        return
    }

    Set-EntryPreviewFromImage -entry $entry -sourceImagePath $imageDialog.FileName -previewSource 'manual' -previewImageUrl '' -listView $listView -tileImageList $tileImageList -detailsBox $detailsBox -statsLabel $statsLabel -logBox $logBox
    $logBox.AppendText("Benutzerbild gesetzt: $($entry.Name)`r`n")
})

$googleImageButton.Add_Click({
    $entry = Get-SelectedEntry -listView $listView
    if (-not $entry) {
        Show-Error 'Bitte zuerst ein Spiel waehlen.'
        return
    }

    $tileImageList = Ensure-TileImageList -listView $listView
    if (-not $tileImageList) {
        Show-Error 'Bildliste konnte nicht initialisiert werden.'
        return
    }

    $queryName = if ([string]::IsNullOrWhiteSpace($entry.Name)) { Split-Path -Leaf $entry.FolderPath } else { $entry.Name }
    $logBox.AppendText("========== Bild-Suche fuer: $queryName ==========`r`n")

    try {
        $queries = @(
            "$queryName DOS game cover",
            "$queryName DOS game screenshot",
            "$queryName retro DOS",
            "$queryName game"
        )

        $imgResult = $null
        
        foreach ($q in $queries) {
            $logBox.AppendText("Query: '$q'`r`n")
            [System.Windows.Forms.Application]::DoEvents()
            
            $imgResult = Get-FirstOnlineImageUrl -query $q
            if ($imgResult -and $imgResult.Url) {
                $logBox.AppendText("  --> Gefunden von: $($imgResult.Source)`r`n")
                break
            } else {
                $logBox.AppendText("  --> Keine Treffer`r`n")
            }
            [System.Windows.Forms.Application]::DoEvents()
        }

        if (-not $imgResult -or -not $imgResult.Url) {
            $logBox.AppendText("FEHLER: Keine Bilder von einer Quelle gefunden`r`n")
            Show-Error ('Kein Online-Bild von Google/Bing/DuckDuckGo/Archive.org f' + $script:umlautu + 'r: ' + $queryName)
            return
        }

        $imgUrl = [string]$imgResult.Url
        $imgSource = [string]$imgResult.Source

        $logBox.AppendText("  URL: $($imgUrl.Substring(0, [Math]::Min(75, $imgUrl.Length)))...`r`n")
        $logBox.AppendText("  Lade herunter...`r`n")
        [System.Windows.Forms.Application]::DoEvents()
        
        Enable-WebTls
        $tempPath = Join-Path $env:TEMP ("dosbox-preview-" + [Guid]::NewGuid().ToString('N') + ".img")
        
        try {
            Invoke-WebRequest -Uri $imgUrl -OutFile $tempPath -UseBasicParsing -TimeoutSec 15 `
                -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
        } catch {
            $logBox.AppendText("  Download fehlgeschlagen: $($_.Exception.Message.Substring(0, [Math]::Min(80, $_.Exception.Message.Length)))`r`n")
            Show-Error "Bild konnte nicht heruntergeladen werden von $imgSource"
            return
        }

        $logBox.AppendText("  Konvertiere zu Vorschau...`r`n")
        [System.Windows.Forms.Application]::DoEvents()

        Set-EntryPreviewFromImage -entry $entry -sourceImagePath $tempPath -previewSource $imgSource -previewImageUrl $imgUrl -listView $listView -tileImageList $tileImageList -detailsBox $detailsBox -statsLabel $statsLabel -logBox $logBox
        $logBox.AppendText("===== OK: Bild von $imgSource gesetzt =====`r`n")

        Remove-Item -LiteralPath $tempPath -ErrorAction SilentlyContinue
    } catch {
        $logBox.AppendText("EXCEPTION: $($_.Exception.Message)`r`n")
        Show-Error ('Fehler bei Bildsuche: ' + $_.Exception.Message)
    }
})

$googleImageBatchButton.Add_Click({
    if (-not $script:Entries -or $script:Entries.Count -eq 0) {
        Show-Error 'Keine Spiele in der Liste vorhanden.'
        return
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
        "WARNUNG: Dieses wird Bilder f$($script:umlautu)r ALLE $($script:Entries.Count) Spiele herunterladen.`r`n`r`nDies kann mehrere Minuten dauern und generiert viele Web-Anfragen.`r`n`r`nFortfahren?",
        'Batch-Download Best' + $script:umlautu + 'tigung',
        'YesNo',
        'Warning'
    )

    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
        $logBox.AppendText("Batch-Download abgebrochen.`r`n")
        return
    }

    $tileImageList = Ensure-TileImageList -listView $listView
    if (-not $tileImageList) {
        Show-Error 'Bildliste konnte nicht initialisiert werden.'
        return
    }

    $logBox.AppendText("`r`n======================== BATCH GOOGLE-BILD DOWNLOAD START ========================`r`n")
    $logBox.AppendText("VORSICHT: Dauer variiert je nach Internetverbindung. Dieses Fenster nicht schlie' + $([char]0x00DF) + 'en!`r`n")
    $logBox.AppendText("Insgesamt: $($script:Entries.Count) Spiele zu verarbeiten`r`n")
    $logBox.AppendText("Startet um: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n")
    $logBox.AppendText("========================================================================================`r`n")
    [System.Windows.Forms.Application]::DoEvents()

    Enable-WebTls
    $successCount = 0
    $failCount = 0
    $skipCount = 0
    $startTime = Get-Date

    foreach ($entry in $script:Entries) {
        try {
            # Skip wenn bereits ein Vorschaubild vorhanden ist
            if ($entry.PreviewImagePath -and (Test-Path $entry.PreviewImagePath)) {
                $logBox.AppendText("[SKIP] $($entry.Name) - Vorschaubild existiert bereits`r`n")
                [System.Windows.Forms.Application]::DoEvents()
                $skipCount++
                continue
            }

            $queryName = if ([string]::IsNullOrWhiteSpace($entry.Name)) { Split-Path -Leaf $entry.FolderPath } else { $entry.Name }
            $logBox.AppendText("[DOWNLOAD $($successCount + $failCount + 1)/$($script:Entries.Count)] $queryName...`r`n")
            [System.Windows.Forms.Application]::DoEvents()

            $queries = @(
                "$queryName DOS game cover",
                "$queryName DOS game screenshot",
                "$queryName retro DOS",
                "$queryName game"
            )

            $imgResult = $null
            foreach ($q in $queries) {
                $imgResult = Get-FirstOnlineImageUrl -query $q
                if ($imgResult -and $imgResult.Url) {
                    break
                }
                [System.Windows.Forms.Application]::DoEvents()
            }

            if (-not $imgResult -or -not $imgResult.Url) {
                $logBox.AppendText("  --> FEHLER: Kein Bild gefunden`r`n")
                [System.Windows.Forms.Application]::DoEvents()
                $failCount++
                continue
            }

            $imgUrl = [string]$imgResult.Url
            $imgSource = [string]$imgResult.Source

            $logBox.AppendText("  --> $imgSource")
            [System.Windows.Forms.Application]::DoEvents()
            
            $tempPath = Join-Path $env:TEMP ("dosbox-preview-" + [Guid]::NewGuid().ToString('N') + ".img")
            
            try {
                Invoke-WebRequest -Uri $imgUrl -OutFile $tempPath -UseBasicParsing -TimeoutSec 15 `
                    -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
                $logBox.AppendText(" - OK`r`n")
            } catch {
                $logBox.AppendText(" - DOWNLOAD FEHLER`r`n")
                $failCount++
                [System.Windows.Forms.Application]::DoEvents()
                continue
            }

            # Konvertiere zu Vorschau
            Set-EntryPreviewFromImage -entry $entry -sourceImagePath $tempPath -previewSource $imgSource -previewImageUrl $imgUrl -listView $listView -tileImageList $tileImageList -detailsBox $detailsBox -statsLabel $statsLabel -logBox $logBox
            
            Remove-Item -LiteralPath $tempPath -ErrorAction SilentlyContinue
            $successCount++
            [System.Windows.Forms.Application]::DoEvents()
        } catch {
            $logBox.AppendText("  --> EXCEPTION: $($_.Exception.Message.Substring(0, [Math]::Min(60, $_.Exception.Message.Length)))`r`n")
            $failCount++
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    $elapsed = $(Get-Date) - $startTime
    $logBox.AppendText("`r`n======================== BATCH GOOGLE-BILD DOWNLOAD FERTIG ========================`r`n")
    $logBox.AppendText("Erfolgreich: $successCount | Fehler: $failCount | Uebersprungen: $skipCount`r`n")
    $logBox.AppendText("Dauer: $([Math]::Floor($elapsed.TotalMinutes))m $($elapsed.Seconds)s`r`n")
    $logBox.AppendText("Beendigung um: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n")
    $logBox.AppendText("========================================================================================`r`n")
    $logBox.AppendText("SPEICHERN NICHT VERGESSEN: 'Cache speichern' klicken!`r`n")
    [System.Windows.Forms.Application]::DoEvents()

    # Refresh Liste
    if ($successCount -gt 0) {
        Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $tileImageList -logBox $logBox -entries $script:Entries
    }
})

$form.Add_Shown({
    # Splitter constraints must be applied after layout; otherwise some systems throw on startup.
    try {
        $split.Panel2MinSize = 520
        $maxLeft = [Math]::Max($split.Panel1MinSize, $split.Width - $split.Panel2MinSize)
        $desiredLeft = [Math]::Min(820, $maxLeft)
        $split.SplitterDistance = [Math]::Max($split.Panel1MinSize, $desiredLeft)
    } catch {
        # Keep startup resilient even if sizing fails on unusual DPI/layout setups.
    }

    $headerToggleButton.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 165), 8)
    $headerToggleButton.BringToFront()
    $rightX = [Math]::Max(12, ($rightPanel.ClientSize.Width - 172))
    $optimizeButton.Location = New-Object System.Drawing.Point($rightX, 310)
    $optionsButton.Location = New-Object System.Drawing.Point($rightX, 354)

    $initialRoot = Load-Settings
    if ($initialRoot) {
        $script:CurrentRoot = $initialRoot
        $pathBox.Text = $script:CurrentRoot
        $logBox.AppendText("Geladener Ordner: $($script:CurrentRoot)`r`n")

        $loaded = Load-FolderCache -rootFolder $script:CurrentRoot
        if ($loaded -ne $null -and $loaded.Count -gt 0) {
            # Initialisiere ImageList und lade Vorschaubilder
            $tileImageList = Ensure-TileImageList -listView $listView
            if ($tileImageList) {
                $script:PreviewsEnabled = $true
                $logBox.AppendText("Lade Vorschaubilder...`r`n")
                [System.Windows.Forms.Application]::DoEvents()
                
                $previewsLoaded = 0
                foreach ($entry in $loaded) {
                    try {
                        $key = Ensure-EntryPreview -entry $entry -imageList $tileImageList -forceRebuild $false -logBox $logBox
                        if ($key -and $key -ne 'default') {
                            $previewsLoaded++
                        }
                    } catch {
                        # Silent continue
                    }
                    [System.Windows.Forms.Application]::DoEvents()
                }
                $logBox.AppendText("Vorschaubilder geladen: $previewsLoaded`r`n")
            }
            
            Populate-List -listView $listView -statsLabel $statsLabel -detailsBox $detailsBox -tileImages $tileImageList -logBox $logBox -entries $loaded
            $logBox.AppendText("Cache geladen: $(Get-CachePath -rootFolder $script:CurrentRoot)`r`n")
        } else {
            $logBox.AppendText("Auto-Scan deaktiviert. Bitte 'Neu scannen' klicken.`r`n")
        }
    } else {
        $logBox.AppendText("Hinweis: Bitte zuerst einen Spieleordner waehlen.`r`n")
    }
})

[void]$form.ShowDialog()
