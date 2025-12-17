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

# -----------------------------
# Helpers
# -----------------------------
function Select-Folder([string]$title) {
    # Prevent dialog from hiding behind other windows
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

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dlg.SelectedPath
    }
    return $null
}

function Select-File([string]$title, [string]$filter, [string]$initialDir) {
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = $title
    $ofd.Filter = $filter
    if ($initialDir -and (Test-Path $initialDir)) { $ofd.InitialDirectory = $initialDir }
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $ofd.FileName
    }
    return $null
}

function Show-ProgressForm {
    param([string]$title)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(620, 170)
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.AutoSize = $false
    $label.Size = New-Object System.Drawing.Size(580, 20)
    $label.Location = New-Object System.Drawing.Point(15, 15)
    $label.Text = "Preparing..."
    $form.Controls.Add($label)

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Minimum = 0
    $bar.Maximum = 100
    $bar.Value = 0
    $bar.Size = New-Object System.Drawing.Size(580, 22)
    $bar.Location = New-Object System.Drawing.Point(15, 45)
    $form.Controls.Add($bar)

    $detail = New-Object System.Windows.Forms.Label
    $detail.AutoSize = $false
    $detail.Size = New-Object System.Drawing.Size(580, 75)
    $detail.Location = New-Object System.Drawing.Point(15, 75)
    $detail.Text = ""
    $form.Controls.Add($detail)

    $form.Show()

    return [pscustomobject]@{
        Form   = $form
        Label  = $label
        Bar    = $bar
        Detail = $detail
    }
}

function Sanitize-FileName([string]$name) {
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) { $name = $name.Replace($c, "_") }
    return $name
}

# Returns $true if Taken Date exists (DateTimeOriginal or CreateDate), else $false
function Has-TakenDate([string]$filePath, [string]$exiftoolPath) {
    try {
        # -s -s -s: values only, one per line
        $out = & $exiftoolPath "-s" "-s" "-s" "-DateTimeOriginal" "-CreateDate" $filePath 2>$null
        if (-not $out) { return $false }

        foreach ($line in $out) {
            $v = ($line | ForEach-Object { $_.ToString().Trim() })
            if ($v -and $v -ne "0000:00:00 00:00:00") {
                return $true
            }
        }
        return $false
    } catch {
        # If probe fails, treat as "missing" so user can force-set it
        return $false
    }
}

# -----------------------------
# Resolve exiftool.exe
# -----------------------------
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
        "exiftool.exe was not found.`n`nExpected locations:`n- .\tools\exiftool\exiftool.exe`n- .\tools\exiftool.exe`n- .\exiftool.exe`n`nPlease place exiftool.exe in one of these paths or select it when prompted.",
        "Missing exiftool",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

# -----------------------------
# 1) Pick source folder
# -----------------------------
$src = Select-Folder "Select Source Folder: photos you want to set the taken date for"
if (-not $src) { exit 0 }

# -----------------------------
# 2) Input date only (YYYY-MM-DD)
# -----------------------------
$defaultDate = (Get-Date).ToString("yyyy-MM-dd")

$dateText = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter the 'Taken Date' ONLY (year-month-day).`nFormat: YYYY-MM-DD`nExample: 2020-01-14`n`nTime will start at 10:00:00 and increase by +1 minute per PROCESSED file.",
    "Enter Taken Date (Date only)",
    $defaultDate
)

if ([string]::IsNullOrWhiteSpace($dateText)) { exit 0 }

try {
    $baseDate = [DateTime]::ParseExact($dateText.Trim(), "yyyy-MM-dd", $null)
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Invalid date format.`n`nUse: YYYY-MM-DD`nExample: 2020-01-14",
        "Invalid Input",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    exit 1
}

# Base time: 10:00:00
$baseDateTime = $baseDate.Date.AddHours(10)

# -----------------------------
# 3) Auto-create destination folder under source
# -----------------------------
$dstName = "_SetTakenDate_Output_" + $baseDate.ToString("yyyyMMdd")
$dst = Join-Path $src $dstName
New-Item -ItemType Directory -Path $dst -Force | Out-Null

# -----------------------------
# 4) Collect files (exclude destination folder to avoid recursion)
# -----------------------------
$exts = @(".jpg",".jpeg",".png",".heic",".tif",".tiff",".cr2",".cr3",".nef",".arw",".dng")

$allFiles = Get-ChildItem -LiteralPath $src -File -Recurse |
    Where-Object {
        $exts -contains $_.Extension.ToLower() -and
        ($_.FullName -notlike ($dst + "\*"))
    } |
    Sort-Object FullName

if (-not $allFiles -or $allFiles.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "No supported files found in the selected source folder.`n`nSupported extensions:`n$($exts -join ', ')",
        "No Files",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
    exit 0
}

# -----------------------------
# 5) Prepare log (in destination folder)
# -----------------------------
$logName = Sanitize-FileName(("SetTakenDate_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date)))
$logPath = Join-Path $dst $logName

"==== Set Taken Date & Copy Log ====" | Out-File -FilePath $logPath -Encoding UTF8
"Time        : $(Get-Date)"           | Out-File -FilePath $logPath -Append -Encoding UTF8
"Source      : $src"                  | Out-File -FilePath $logPath -Append -Encoding UTF8
"Destination : $dst"                  | Out-File -FilePath $logPath -Append -Encoding UTF8
"Taken Date  : $($baseDate.ToString('yyyy-MM-dd'))" | Out-File -FilePath $logPath -Append -Encoding UTF8
"Rule        : Start 10:00:00, +1 minute per PROCESSED file (sorted by FullName)" | Out-File -FilePath $logPath -Append -Encoding UTF8
"Process     : ONLY files missing DateTimeOriginal/CreateDate" | Out-File -FilePath $logPath -Append -Encoding UTF8
"ExifTool    : $exiftool"             | Out-File -FilePath $logPath -Append -Encoding UTF8
"Total Files : $($allFiles.Count)"    | Out-File -FilePath $logPath -Append -Encoding UTF8
"-----------------------------------" | Out-File -FilePath $logPath -Append -Encoding UTF8

# -----------------------------
# 6) Scan metadata to filter targets
# -----------------------------
$ui = Show-ProgressForm -title "Scanning: Detect Missing Taken Date"
$targets = New-Object System.Collections.Generic.List[System.IO.FileInfo]
$skipped = 0

for ($i = 0; $i -lt $allFiles.Count; $i++) {
    $f = $allFiles[$i]
    $percent = [int](($i + 1) / $allFiles.Count * 100)

    $ui.Label.Text = "Scanning: $($i+1) / $($allFiles.Count)  ($percent%)"
    $ui.Bar.Value = [Math]::Min(100, [Math]::Max(0, $percent))
    $ui.Detail.Text = $f.FullName
    $ui.Form.Refresh()

    $has = Has-TakenDate -filePath $f.FullName -exiftoolPath $exiftool
    if (-not $has) {
        $targets.Add($f) | Out-Null
    } else {
        $skipped++
        "[SKIP] Has TakenDate | $($f.FullName)" | Out-File -FilePath $logPath -Append -Encoding UTF8
    }

    [System.Windows.Forms.Application]::DoEvents() | Out-Null
}

$ui.Form.Close()

if ($targets.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "All files already have a Taken Date (DateTimeOriginal/CreateDate).`nNothing to process.",
        "Nothing to do",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    Start-Process "explorer.exe" -ArgumentList "`"$dst`""
    exit 0
}

# -----------------------------
# 7) Process ONLY targets
# -----------------------------
$ui = Show-ProgressForm -title "Processing: Copy + Set Taken Date (Missing only)"
$ok = 0
$fail = 0

for ($i = 0; $i -lt $targets.Count; $i++) {
    $f = $targets[$i]
    $percent = [int](($i + 1) / $targets.Count * 100)

    # Timestamp: base 10:00:00 + i minutes (ONLY for processed files)
    $dt = $baseDateTime.AddMinutes($i)
    $exifDate = $dt.ToString("yyyy:MM:dd HH:mm:ss")

    $ui.Label.Text = "Processing: $($i+1) / $($targets.Count)  ($percent%)"
    $ui.Bar.Value = [Math]::Min(100, [Math]::Max(0, $percent))
    $ui.Detail.Text = "TakenDate: $($dt.ToString('yyyy-MM-dd HH:mm:ss'))`nFile: $($f.FullName)"
    $ui.Form.Refresh()

    # Keep relative structure under destination
    $rel = $f.FullName.Substring($src.Length).TrimStart('\','/')
    $targetPath = Join-Path $dst $rel
    $targetDir  = Split-Path -Parent $targetPath
    if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }

    try {
        # Copy first (do NOT touch original)
        Copy-Item -LiteralPath $f.FullName -Destination $targetPath -Force

        # exiftool argument splitting fix:
        # Use call operator (&) with splatted array so each item stays as one argument.
        $args = @(
            "-overwrite_original",
            "-P",
            "-m",
            ("-AllDates=$exifDate"),
            ("-FileModifyDate=$exifDate"),
            ("-FileCreateDate=$exifDate"),
            "-charset", "filename=UTF8",
            $targetPath
        )

        & $exiftool @args | Out-Null
        $exit = $LASTEXITCODE

        if ($exit -eq 0) {
            $ok++
            "[OK]   $($dt.ToString('yyyy-MM-dd HH:mm:ss')) | $($f.FullName) -> $targetPath" |
                Out-File -FilePath $logPath -Append -Encoding UTF8
        } else {
            $fail++
            "[FAIL] ExitCode=$exit | $($dt.ToString('yyyy-MM-dd HH:mm:ss')) | $targetPath" |
                Out-File -FilePath $logPath -Append -Encoding UTF8
        }
    }
    catch {
        $fail++
        "[FAIL] Exception | $($dt.ToString('yyyy-MM-dd HH:mm:ss')) | $($f.FullName) | $($_.Exception.Message)" |
            Out-File -FilePath $logPath -Append -Encoding UTF8
    }

    [System.Windows.Forms.Application]::DoEvents() | Out-Null
}

$ui.Form.Close()

# -----------------------------
# 8) Result popup + open destination
# -----------------------------
"-----------------------------------" | Out-File -FilePath $logPath -Append -Encoding UTF8
"Scan Result : Total=$($allFiles.Count), Targets=$($targets.Count), Skipped=$skipped" | Out-File -FilePath $logPath -Append -Encoding UTF8
"Done. OK=$ok, FAIL=$fail"            | Out-File -FilePath $logPath -Append -Encoding UTF8

$msg = "Completed.`n`nTotal found: $($allFiles.Count)`nProcessed (missing only): $($targets.Count)`nSkipped: $skipped`n`nSuccess: $ok`nFailed: $fail`n`nDestination folder:`n$dst`n`nLog file:`n$logPath"
[System.Windows.Forms.MessageBox]::Show(
    $msg,
    "Result",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null

Start-Process "explorer.exe" -ArgumentList "`"$dst`""
