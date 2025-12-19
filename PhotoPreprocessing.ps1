#requires -version 5.1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---- Ensure STA (WinForms stability) ----
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process powershell -ArgumentList @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-STA",
        "-File", "`"$PSCommandPath`""
    )
    exit
}

# =========================
# UI helpers
# =========================
function Select-Folder([string]$title) {
    $owner = New-Object System.Windows.Forms.Form
    $owner.Text = ""
    $owner.StartPosition = "CenterScreen"
    $owner.Size = New-Object System.Drawing.Size(1,1)
    $owner.TopMost = $true
    $owner.ShowInTaskbar = $false
    $owner.Opacity = 0
    $owner.Show()
    $owner.Activate()

    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $title
    $dlg.ShowNewFolderButton = $true

    $result = $dlg.ShowDialog($owner)

    $owner.Close()
    $owner.Dispose()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
    return $null
}

function Select-File([string]$title, [string]$filter, [string]$initialDir) {
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = $title
    $ofd.Filter = $filter
    if ($initialDir -and (Test-Path $initialDir)) { $ofd.InitialDirectory = $initialDir }
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $ofd.FileName }
    return $null
}

function Show-ProgressForm {
    param([string]$title)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(760, 200)
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.AutoSize = $false
    $label.Size = New-Object System.Drawing.Size(720, 20)
    $label.Location = New-Object System.Drawing.Point(15, 15)
    $label.Text = "Preparing..."
    $form.Controls.Add($label)

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Minimum = 0
    $bar.Maximum = 100
    $bar.Value = 0
    $bar.Size = New-Object System.Drawing.Size(720, 22)
    $bar.Location = New-Object System.Drawing.Point(15, 45)
    $form.Controls.Add($bar)

    $detail = New-Object System.Windows.Forms.Label
    $detail.AutoSize = $false
    $detail.Size = New-Object System.Drawing.Size(720, 110)
    $detail.Location = New-Object System.Drawing.Point(15, 75)
    $detail.Text = ""
    $form.Controls.Add($detail)

    $form.Show()

    return [pscustomobject]@{ Form=$form; Label=$label; Bar=$bar; Detail=$detail }
}

function Sanitize-FileName([string]$name) {
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) { $name = $name.Replace($c, "_") }
    return $name
}

# =========================
# File type / extensions
# =========================
$PHOTO_EXTS = @(".jpg",".jpeg",".png",".heic",".tif",".tiff",".cr2",".cr3",".nef",".arw",".dng")
$MOVIE_EXTS = @(".mp4",".mov",".m4v",".avi",".wmv",".mts",".m2ts")
$SUPPORTED_EXTS = @($PHOTO_EXTS + $MOVIE_EXTS)

function Get-ContentType([string]$ext) {
    $e = $ext.ToLower()
    if ($PHOTO_EXTS -contains $e) { return "PHOTO" }
    if ($MOVIE_EXTS -contains $e) { return "MOVIE" }
    return $null
}

# =========================
# Folder date (YYYY-MM-DD)
# =========================
function Get-DateFromFolderPath {
    param([string]$fileFullName, [string]$srcRoot)

    try {
        $dir = Split-Path -Parent $fileFullName
        $root = (Resolve-Path -LiteralPath $srcRoot).Path.TrimEnd('\')

        while ($dir -and ($dir.Length -ge $root.Length)) {
            $leaf = Split-Path -Leaf $dir
            if ($leaf -match '^(?<y>\d{4})-(?<m>\d{2})-(?<d>\d{2})$') {
                $s = $matches[0]
                try { return [DateTime]::ParseExact($s, "yyyy-MM-dd", $null) } catch { }
            }
            if ($dir -ieq $root) { break }
            $dir = Split-Path -Parent $dir
        }
    } catch { }
    return $null
}

# =========================
# IMG naming rules
# =========================
function Get-Hash4FromNameSize {
    param([string]$nameNoExt, [long]$sizeBytes)
    $s = "{0}|{1}" -f $nameNoExt, $sizeBytes
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($s)
    $hash = $sha1.ComputeHash($bytes)
    $n = [BitConverter]::ToUInt32($hash, 0)
    $num = [int]($n % 10000)
    return $num
}

function Get-BaseImgNumber {
    param([string]$baseNameNoExt, [long]$sizeBytes)

    $m1 = [regex]::Match($baseNameNoExt, '^(?i)IMG[_-]?(?<n>\d{4})$')
    if ($m1.Success) { return [int]$m1.Groups["n"].Value }

    $m2 = [regex]::Match($baseNameNoExt, '^(?i)Law_.*_(?<n>\d{4})$')
    if ($m2.Success) { return [int]$m2.Groups["n"].Value }

    return (Get-Hash4FromNameSize -nameNoExt $baseNameNoExt -sizeBytes $sizeBytes)
}

function Resolve-UniqueImgFileName {
    param(
        [int]$startNum,
        [string]$ext,
        [string]$destDir,
        [hashtable]$usedMap
    )

    $dirKey = $destDir.ToLower()
    if (-not $usedMap.ContainsKey($dirKey)) {
        $usedMap[$dirKey] = New-Object System.Collections.Generic.HashSet[string]
        if (Test-Path $destDir) {
            Get-ChildItem -LiteralPath $destDir -File -ErrorAction SilentlyContinue | ForEach-Object {
                [void]$usedMap[$dirKey].Add($_.Name.ToLower())
            }
        }
    }

    $num = $startNum
    for ($i=0; $i -lt 11000; $i++) {
        $name = ("IMG_{0:0000}{1}" -f $num, $ext.ToLower())
        $lower = $name.ToLower()
        if (-not $usedMap[$dirKey].Contains($lower)) {
            [void]$usedMap[$dirKey].Add($lower)
            return $name
        }
        $num = ($num + 1) % 10000
    }

    $fallback = ("IMG_{0:0000}_{1}{2}" -f $startNum, ([Guid]::NewGuid().ToString("N").Substring(0,4)), $ext.ToLower())
    [void]$usedMap[$dirKey].Add($fallback.ToLower())
    return $fallback
}

# =========================
# exiftool resolve
# =========================
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$candidates = @(
    (Join-Path $scriptDir "tools\exiftool\exiftool.exe"),
    (Join-Path $scriptDir "tools\exiftool.exe"),
    (Join-Path $scriptDir "exiftool.exe"),
    (Join-Path $scriptDir "tools\exiftool\exiftool(-k).exe"),
    (Join-Path $scriptDir "tools\exiftool(-k).exe")
)
$exiftool = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $exiftool) {
    $picked = Select-File -title "Select exiftool.exe (not found automatically)" -filter "EXE (*.exe)|*.exe" -initialDir $scriptDir
    if ($picked) { $exiftool = $picked }
}
if (-not $exiftool -or -not (Test-Path $exiftool)) {
    [System.Windows.Forms.MessageBox]::Show(
        "exiftool.exe was not found. Place it under tools\exiftool\exiftool.exe or select it.",
        "Missing exiftool",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

# =========================
# FAST metadata scan (single exiftool call)
# =========================
$READ_TAGS = @(
    "-DateTimeOriginal",
    "-CreateDate",
    "-QuickTime:CreateDate",
    "-QuickTime:MediaCreateDate"
)

function Get-MetaMapBatch {
    param(
        [string]$exiftoolPath,
        [System.IO.FileInfo[]]$files,
        [string]$tempDir
    )

    $map = @{}

    $listPath = Join-Path $tempDir ("exif_list_{0}.txt" -f ([Guid]::NewGuid().ToString("N")))
    $jsonPath = Join-Path $tempDir ("exif_out_{0}.json" -f ([Guid]::NewGuid().ToString("N")))

    try {
        $files | ForEach-Object { $_.FullName } | Out-File -FilePath $listPath -Encoding UTF8

        # ✅ (중요) filename charset 지정 (한글/공백 파일 안정)
        $args = @(
            "-charset","filename=UTF8",
            "-json","-G1","-a","-s"
        ) + $READ_TAGS + @("-@", $listPath)

        $raw = & $exiftoolPath @args 2>$null
        if (-not $raw) { return $map }

        $raw | Out-File -FilePath $jsonPath -Encoding UTF8
        $items = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json

        foreach ($it in $items) {
            $src = $it.SourceFile
            if (-not $src) { continue }

            $cands = @(
                "EXIF:DateTimeOriginal",
                "EXIF:CreateDate",
                "QuickTime:CreateDate",
                "QuickTime:MediaCreateDate",
                "DateTimeOriginal",
                "CreateDate"
            )

            $bestRaw = $null
            $bestTag = $null

            foreach ($p in $cands) {
                if ($it.PSObject.Properties.Name -contains $p) {
                    $v = ($it.$p | ForEach-Object { $_.ToString().Trim() })
                    if ($v -and $v -ne "0000:00:00 00:00:00") {
                        $bestRaw = $v
                        $bestTag = $p
                        break
                    }
                }
            }

            $dateOnly = $null
            if ($bestRaw) {
                $m = [regex]::Match($bestRaw, '^(?<y>\d{4})[:\-](?<m>\d{2})[:\-](?<d>\d{2})')
                if ($m.Success) {
                    $dateOnly = "{0}-{1}-{2}" -f $m.Groups["y"].Value, $m.Groups["m"].Value, $m.Groups["d"].Value
                }
            }

            $map[$src] = [pscustomobject]@{
                RawBest  = $bestRaw
                DateOnly = $dateOnly
                FoundTag = $bestTag
            }
        }
    } finally {
        if (Test-Path $listPath) { Remove-Item $listPath -Force -ErrorAction SilentlyContinue }
        if (Test-Path $jsonPath) { Remove-Item $jsonPath -Force -ErrorAction SilentlyContinue }
    }

    return $map
}

# =========================
# FINAL SAFETY CHECK (last method)
# - Check on the actual file we are about to write (destPath)
# - If taken date exists, NEVER overwrite
# =========================
function Has-TakenDateFast {
    param([string]$exiftoolPath, [string]$filePath)

    try {
        $args = @(
            "-charset","filename=UTF8",
            "-s","-s","-s",
            "-DateTimeOriginal",
            "-CreateDate",
            "-QuickTime:CreateDate",
            "-QuickTime:MediaCreateDate",
            $filePath
        )
        $out = & $exiftoolPath @args 2>$null
        foreach ($line in $out) {
            $v = $line.ToString().Trim()
            if ($v -and $v -ne "0000:00:00 00:00:00") { return $true }
        }
    } catch { }
    return $false
}

# =========================
# Write metadata (type-aware)
# =========================
function Exif-SetTakenDate {
    param(
        [string]$exiftoolPath,
        [string]$contentType,
        [string]$targetPath,
        [DateTime]$dt
    )
    $exifDate = $dt.ToString("yyyy:MM:dd HH:mm:ss")

    $baseArgs = @(
        "-overwrite_original",
        "-P",
        "-m",
        "-charset","filename=UTF8"
    )

    if ($contentType -eq "MOVIE") {
        $setArgs = @(
            ("-QuickTime:CreateDate=$exifDate"),
            ("-QuickTime:MediaCreateDate=$exifDate"),
            ("-MediaCreateDate=$exifDate"),
            ("-TrackCreateDate=$exifDate"),
            ("-CreateDate=$exifDate"),
            ("-ModifyDate=$exifDate"),
            ("-FileModifyDate=$exifDate"),
            ("-FileCreateDate=$exifDate")
        )
    } else {
        $setArgs = @(
            ("-AllDates=$exifDate"),
            ("-FileModifyDate=$exifDate"),
            ("-FileCreateDate=$exifDate")
        )
    }

    $args = @() + $baseArgs + $setArgs + @($targetPath)
    & $exiftoolPath @args | Out-Null
    return $LASTEXITCODE
}

# =========================
# Pick source / output
# =========================
$src = Select-Folder "Select Source Folder (root): contains date folders like YYYY-MM-DD"
if (-not $src) { exit 0 }
$src = (Resolve-Path -LiteralPath $src).Path.TrimEnd('\')

$outRoot = Select-Folder "Select Output Folder (root): result will be saved under PHOTO/MOVIE/ERROR"
if (-not $outRoot) { exit 0 }
$outRoot = (Resolve-Path -LiteralPath $outRoot).Path.TrimEnd('\')

$photoRoot = Join-Path $outRoot "PHOTO"
$movieRoot = Join-Path $outRoot "MOVIE"
$errorRoot = Join-Path $outRoot "ERROR"
$logRoot   = Join-Path $outRoot "LOG"
New-Item -ItemType Directory -Path $photoRoot -Force | Out-Null
New-Item -ItemType Directory -Path $movieRoot -Force | Out-Null
New-Item -ItemType Directory -Path $errorRoot -Force | Out-Null
New-Item -ItemType Directory -Path $logRoot   -Force | Out-Null

$runId = (Get-Date).ToString("yyyyMMdd_HHmmss")
$logPath = Join-Path $logRoot (Sanitize-FileName("PhotoPreprocess_$runId.log"))

"==== Photo Preprocessing Log ====" | Out-File -FilePath $logPath -Encoding UTF8
"RunTime    : $(Get-Date)"          | Out-File -FilePath $logPath -Append -Encoding UTF8
"SourceRoot : $src"                 | Out-File -FilePath $logPath -Append -Encoding UTF8
"OutputRoot : $outRoot"             | Out-File -FilePath $logPath -Append -Encoding UTF8
"ExifTool   : $exiftool"            | Out-File -FilePath $logPath -Append -Encoding UTF8
"---------------------------------" | Out-File -FilePath $logPath -Append -Encoding UTF8

# =========================
# Collect files
# =========================
$ui = Show-ProgressForm -title "Scanning: collecting files"
$allFiles = Get-ChildItem -LiteralPath $src -File -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $SUPPORTED_EXTS -contains $_.Extension.ToLower() } |
    Sort-Object FullName

if (-not $allFiles -or $allFiles.Count -eq 0) {
    $ui.Form.Close()
    [System.Windows.Forms.MessageBox]::Show(
        "No supported files found under source folder.",
        "No Files",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
    exit 0
}
$ui.Detail.Text = "Files: $($allFiles.Count)`nReading metadata (batch)..."
$ui.Form.Refresh()

# =========================
# Batch read metadata once (speed)
# =========================
$tempDir = Join-Path $env:TEMP "PhotoPreprocess"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$metaMap = Get-MetaMapBatch -exiftoolPath $exiftool -files $allFiles -tempDir $tempDir

$ui.Form.Close()

# =========================
# Build records + targets
# =========================
$ui = Show-ProgressForm -title "Scanning: building plan"
$records = New-Object System.Collections.Generic.List[object]
$targets = New-Object System.Collections.Generic.List[object]

$cntNoDateFolder = 0
$cntKeepAll = 0
$cntTarget = 0

for ($i=0; $i -lt $allFiles.Count; $i++) {
    $f = $allFiles[$i]
    $pct = [int](($i+1)/$allFiles.Count*100)
    $ui.Label.Text = "Scanning: $($i+1) / $($allFiles.Count) ($pct%)"
    $ui.Bar.Value = [Math]::Min(100, [Math]::Max(0, $pct))
    $ui.Detail.Text = $f.FullName
    $ui.Form.Refresh()
    [System.Windows.Forms.Application]::DoEvents() | Out-Null

    $ctype = Get-ContentType $f.Extension
    if (-not $ctype) { continue }

    $folderDate = Get-DateFromFolderPath -fileFullName $f.FullName -srcRoot $src

    $meta = $null
    if ($metaMap.ContainsKey($f.FullName)) { $meta = $metaMap[$f.FullName] }
    $hasMetaDate = ($meta -and $meta.RawBest)
    $metaDateOnly = if ($meta) { $meta.DateOnly } else { $null }

    if (-not $folderDate) {
        $cntNoDateFolder++
        $rec = [pscustomobject]@{
            FileInfo      = $f
            ContentType   = $ctype
            FolderDate    = $null
            HasMetaDate   = $hasMetaDate
            MetaDateOnly  = $metaDateOnly
            IsTarget      = $false
            Reason        = "NO_DATE_FOLDER"
        }
        $records.Add($rec) | Out-Null
        continue
    }

    # ✅ 정책: 메타가 있으면 절대 변경하지 않음 / 없을 때만 세팅
    $isTarget = (-not $hasMetaDate)
    $reason = if ($isTarget) { "MISSING_TAKEN_DATE" } else { "KEEP_META" }

    $rec2 = [pscustomobject]@{
        FileInfo      = $f
        ContentType   = $ctype
        FolderDate    = $folderDate
        HasMetaDate   = $hasMetaDate
        MetaDateOnly  = $metaDateOnly
        IsTarget      = $isTarget
        Reason        = $reason
    }
    $records.Add($rec2) | Out-Null

    if ($isTarget) { $targets.Add($rec2) | Out-Null; $cntTarget++ }
    else { $cntKeepAll++ }
}
$ui.Form.Close()

# targets-only sorted order (deterministic dt assignment)
$targetsSorted = $targets | Sort-Object { $_.FileInfo.FullName }
$targetIndexMap = @{}
for ($k=0; $k -lt $targetsSorted.Count; $k++) {
    $targetIndexMap[$targetsSorted[$k].FileInfo.FullName] = $k
}

"TotalFiles   : $($allFiles.Count)" | Out-File -FilePath $logPath -Append -Encoding UTF8
"Targets(meta): $cntTarget"         | Out-File -FilePath $logPath -Append -Encoding UTF8
"Keep(meta)   : $cntKeepAll"        | Out-File -FilePath $logPath -Append -Encoding UTF8
"NoDateFolder : $cntNoDateFolder"   | Out-File -FilePath $logPath -Append -Encoding UTF8
"---------------------------------" | Out-File -FilePath $logPath -Append -Encoding UTF8

# =========================
# Process (copy + rename + maybe set meta)
# =========================
$ui = Show-ProgressForm -title "Processing: copy + rename + metadata"
$usedMap = @{}

$UI_EVERY = 2

$ok = 0; $fail = 0
$cntRenameOnly = 0
$cntMetaOnly = 0
$cntRenameAndMeta = 0
$cntKeepNoChange = 0
$cntErrNoDateFolder = 0
$cntErrProcess = 0
$cntMetaSkippedBySafeguard = 0

for ($i=0; $i -lt $records.Count; $i++) {
    $rec = $records[$i]
    $f = $rec.FileInfo

    if ( ($i % $UI_EVERY) -eq 0 -or $i -eq ($records.Count - 1) ) {
        $pct = [int](($i+1)/$records.Count*100)
        $ui.Label.Text = "Processing: $($i+1)/$($records.Count) ($pct%) | OK=$ok FAIL=$fail"
        $ui.Bar.Value  = [Math]::Min(100, [Math]::Max(0, $pct))
        $ui.Detail.Text = "Reason: $($rec.Reason)`nFile: $($f.FullName)"
        $ui.Form.Refresh()
        [System.Windows.Forms.Application]::DoEvents() | Out-Null
    }

    try {
        $ext = $f.Extension.ToLower()
        $origName = $f.Name

        if ($rec.Reason -eq "NO_DATE_FOLDER") {
            $rel = $f.FullName.Substring($src.Length).TrimStart('\','/')
            $dest = Join-Path (Join-Path $errorRoot "NoDateFolder") $rel
            $destDir2 = Split-Path -Parent $dest
            New-Item -ItemType Directory -Path $destDir2 -Force | Out-Null
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force
            $cntErrNoDateFolder++
            "[ERROR] NoDateFolder | $($f.FullName) -> $dest" | Out-File -FilePath $logPath -Append -Encoding UTF8
            continue
        }

        # ---- Decide classification date ----
        # meta 유지 파일이면 metaDateOnly 기준으로 폴더 분류(있을 때)
        $classDate = $rec.FolderDate
        if ($rec.Reason -eq "KEEP_META" -and $rec.MetaDateOnly) {
            try { $classDate = [DateTime]::ParseExact($rec.MetaDateOnly, "yyyy-MM-dd", $null) } catch { $classDate = $rec.FolderDate }
        }

        $yyyymmdd = $classDate.ToString("yyyy-MM-dd")
        $contentRoot = if ($rec.ContentType -eq "MOVIE") { $movieRoot } else { $photoRoot }
        $destDir = Join-Path $contentRoot $yyyymmdd
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null

        # ---- Rename ----
        $nameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $num = Get-BaseImgNumber -baseNameNoExt $nameNoExt -sizeBytes $f.Length
        $finalName = Resolve-UniqueImgFileName -startNum $num -ext $ext -destDir $destDir -usedMap $usedMap
        $destPath = Join-Path $destDir $finalName

        Copy-Item -LiteralPath $f.FullName -Destination $destPath -Force

        $renameChanged = ($origName.ToLower() -ne $finalName.ToLower())

        # ---- Metadata set (targets only) ----
        $metaChanged = $false
        if ($rec.IsTarget) {

            # ✅✅✅ 마지막 방법: "쓰기 직전" 실물 파일 기준 최종 안전장치
            if (Has-TakenDateFast -exiftoolPath $exiftool -filePath $destPath) {
                $cntMetaSkippedBySafeguard++
                "[SAFEGUARD] Skip meta write (already exists) | $destPath" |
                    Out-File -FilePath $logPath -Append -Encoding UTF8
            }
            else {
                $tIndex = $targetIndexMap[$f.FullName]
                $dt = $rec.FolderDate.Date.AddHours(9).AddSeconds([double]$tIndex)  # auto roll-over
                $exit = Exif-SetTakenDate -exiftoolPath $exiftool -contentType $rec.ContentType -targetPath $destPath -dt $dt

                if ($exit -ne 0) {
                    $fail++; $cntErrProcess++
                    "[FAIL] ExifTool ExitCode=$exit | $($f.FullName) -> $destPath" | Out-File -FilePath $logPath -Append -Encoding UTF8

                    $errDir = Join-Path (Join-Path $errorRoot "ExifFail") $yyyymmdd
                    New-Item -ItemType Directory -Path $errDir -Force | Out-Null
                    Copy-Item -LiteralPath $destPath -Destination (Join-Path $errDir $finalName) -Force
                    continue
                }

                $metaChanged = $true
                "[SET] MissingMeta -> $($rec.FolderDate.ToString('yyyy-MM-dd')) | $($f.FullName) -> $destPath" |
                    Out-File -FilePath $logPath -Append -Encoding UTF8
            }
        } else {
            "[COPY] KeepMeta | $($f.FullName) -> $destPath" | Out-File -FilePath $logPath -Append -Encoding UTF8
        }

        if ($renameChanged -and $metaChanged) { $cntRenameAndMeta++ }
        elseif ($renameChanged -and -not $metaChanged) { $cntRenameOnly++ }
        elseif (-not $renameChanged -and $metaChanged) { $cntMetaOnly++ }
        else { $cntKeepNoChange++ }

        $ok++
    }
    catch {
        $fail++; $cntErrProcess++
        $em = $_.Exception.Message
        "[FAIL] Exception | $($f.FullName) | $em" | Out-File -FilePath $logPath -Append -Encoding UTF8

        try {
            $rel2 = $f.FullName.Substring($src.Length).TrimStart('\','/')
            $destErr = Join-Path (Join-Path $errorRoot "Failed") $rel2
            $destErrDir = Split-Path -Parent $destErr
            New-Item -ItemType Directory -Path $destErrDir -Force | Out-Null
            Copy-Item -LiteralPath $f.FullName -Destination $destErr -Force
            "[ERROR] Copied(Failed) | $($f.FullName) -> $destErr" | Out-File -FilePath $logPath -Append -Encoding UTF8
        } catch { }
        continue
    }
}

$ui.Form.Close()

# =========================
# Result popup
# =========================
"---------------------------------" | Out-File -FilePath $logPath -Append -Encoding UTF8
"Summary: OK=$ok FAIL=$fail"        | Out-File -FilePath $logPath -Append -Encoding UTF8
"RenameOnly     : $cntRenameOnly"   | Out-File -FilePath $logPath -Append -Encoding UTF8
"MetaOnly       : $cntMetaOnly"     | Out-File -FilePath $logPath -Append -Encoding UTF8
"Rename+Meta    : $cntRenameAndMeta"| Out-File -FilePath $logPath -Append -Encoding UTF8
"KeepNoChange   : $cntKeepNoChange" | Out-File -FilePath $logPath -Append -Encoding UTF8
"MetaSkip(Safeguard): $cntMetaSkippedBySafeguard" | Out-File -FilePath $logPath -Append -Encoding UTF8
"NoDateFolder->ERROR : $cntErrNoDateFolder" | Out-File -FilePath $logPath -Append -Encoding UTF8
"ProcessError->ERROR : $cntErrProcess"      | Out-File -FilePath $logPath -Append -Encoding UTF8

$msg = @"
Completed.

Total files found: $($allFiles.Count)
Targets(meta set): $cntTarget
Keep(meta as-is):  $cntKeepAll
No date folder:    $cntNoDateFolder

--- Changes breakdown ---
Rename only      : $cntRenameOnly
Metadata only    : $cntMetaOnly
Rename + Metadata: $cntRenameAndMeta
No change        : $cntKeepNoChange
Meta skipped by safeguard (already had date): $cntMetaSkippedBySafeguard

--- Errors ---
NoDateFolder -> ERROR : $cntErrNoDateFolder
Processing  -> ERROR : $cntErrProcess

Success: $ok
Failed : $fail

Output root:
$outRoot

Log:
$logPath
"@

[System.Windows.Forms.MessageBox]::Show(
    $msg,
    "Photo Preprocessing Result",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null

Start-Process "explorer.exe" -ArgumentList "`"$outRoot`""
