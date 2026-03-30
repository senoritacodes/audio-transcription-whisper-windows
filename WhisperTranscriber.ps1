Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:NativeWhisperAudioExtensions = @(
    ".wav", ".mp3", ".ogg", ".flac"
)
$script:ConvertibleAudioExtensions = @(
    ".m4a", ".aac", ".wma", ".opus", ".mp4", ".m4b"
)
$script:SupportedAudioExtensions = @($script:NativeWhisperAudioExtensions + $script:ConvertibleAudioExtensions | Select-Object -Unique)
$script:SelectedAudioFiles = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

function Get-WhisperLanguageOptions {
    $pairs = @(
        "auto|Auto Detect",
        "af|Afrikaans",
        "am|Amharic",
        "ar|Arabic",
        "as|Assamese",
        "az|Azerbaijani",
        "ba|Bashkir",
        "be|Belarusian",
        "bg|Bulgarian",
        "bn|Bengali",
        "bo|Tibetan",
        "br|Breton",
        "bs|Bosnian",
        "ca|Catalan",
        "cs|Czech",
        "cy|Welsh",
        "da|Danish",
        "de|German",
        "el|Greek",
        "en|English",
        "es|Spanish",
        "et|Estonian",
        "eu|Basque",
        "fa|Persian",
        "fi|Finnish",
        "fo|Faroese",
        "fr|French",
        "gl|Galician",
        "gu|Gujarati",
        "ha|Hausa",
        "haw|Hawaiian",
        "he|Hebrew",
        "hi|Hindi",
        "hr|Croatian",
        "ht|Haitian Creole",
        "hu|Hungarian",
        "hy|Armenian",
        "id|Indonesian",
        "is|Icelandic",
        "it|Italian",
        "ja|Japanese",
        "jw|Javanese",
        "ka|Georgian",
        "kk|Kazakh",
        "km|Khmer",
        "kn|Kannada",
        "ko|Korean",
        "la|Latin",
        "lb|Luxembourgish",
        "ln|Lingala",
        "lo|Lao",
        "lt|Lithuanian",
        "lv|Latvian",
        "mg|Malagasy",
        "mi|Maori",
        "mk|Macedonian",
        "ml|Malayalam",
        "mn|Mongolian",
        "mr|Marathi",
        "ms|Malay",
        "mt|Maltese",
        "my|Burmese",
        "ne|Nepali",
        "nl|Dutch",
        "nn|Norwegian Nynorsk",
        "no|Norwegian",
        "oc|Occitan",
        "pa|Punjabi",
        "pl|Polish",
        "ps|Pashto",
        "pt|Portuguese",
        "ro|Romanian",
        "ru|Russian",
        "sa|Sanskrit",
        "sd|Sindhi",
        "si|Sinhala",
        "sk|Slovak",
        "sl|Slovenian",
        "sn|Shona",
        "so|Somali",
        "sq|Albanian",
        "sr|Serbian",
        "su|Sundanese",
        "sv|Swedish",
        "sw|Swahili",
        "ta|Tamil",
        "te|Telugu",
        "tg|Tajik",
        "th|Thai",
        "tk|Turkmen",
        "tl|Tagalog",
        "tr|Turkish",
        "tt|Tatar",
        "uk|Ukrainian",
        "ur|Urdu",
        "uz|Uzbek",
        "vi|Vietnamese",
        "yi|Yiddish",
        "yo|Yoruba",
        "zh|Chinese"
    )

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($pair in $pairs) {
        $parts = $pair.Split("|", 2)
        $result.Add([PSCustomObject]@{
            Code  = $parts[0]
            Label = ("{0} ({1})" -f $parts[1], $parts[0])
        })
    }

    return $result.ToArray()
}

function New-Label {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 120
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, 20)
    return $label
}

function New-TextBox {
    param(
        [int]$X,
        [int]$Y,
        [int]$Width = 600
    )

    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point($X, $Y)
    $box.Size = New-Object System.Drawing.Size($Width, 24)
    return $box
}

function Add-Log {
    param(
        [System.Windows.Forms.TextBoxBase]$LogTextBox,
        [string]$Message
    )

    $time = Get-Date -Format "HH:mm:ss"
    $LogTextBox.AppendText("[$time] $Message`r`n")
}

function Format-ArgsForLog {
    param([string[]]$Arguments)

    $parts = foreach ($arg in $Arguments) {
        if ($null -eq $arg) {
            continue
        }
        if ($arg -match '\s') {
            '"{0}"' -f $arg.Replace('"', '\"')
        }
        else {
            $arg
        }
    }

    return ($parts -join " ")
}

function ConvertTo-ProcessArgumentString {
    param([string[]]$Arguments)

    $parts = foreach ($arg in $Arguments) {
        if ($null -eq $arg) {
            continue
        }

        if ($arg.Length -eq 0) {
            '""'
            continue
        }

        if ($arg -notmatch '[\s"]') {
            $arg
            continue
        }

        $escaped = $arg -replace '(\\*)"', '$1$1\"'
        $escaped = $escaped -replace '(\\+)$', '$1$1'
        '"' + $escaped + '"'
    }

    return ($parts -join " ")
}

function Get-FFmpegPath {
    param([string]$WhisperExePath)

    try {
        $ffmpegCmd = Get-Command ffmpeg -ErrorAction Stop
        if ($ffmpegCmd -and -not [string]::IsNullOrWhiteSpace($ffmpegCmd.Source)) {
            return $ffmpegCmd.Source
        }
    }
    catch {
    }

    if (-not [string]::IsNullOrWhiteSpace($WhisperExePath) -and (Test-Path -LiteralPath $WhisperExePath -PathType Leaf)) {
        $exeDir = [System.IO.Path]::GetDirectoryName($WhisperExePath)
        if (-not [string]::IsNullOrWhiteSpace($exeDir)) {
            $candidate = Join-Path -Path $exeDir -ChildPath "ffmpeg.exe"
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return $candidate
            }
        }
    }

    return $null
}

function Install-FFmpegWithWinget {
    try {
        $winget = Get-Command winget -ErrorAction Stop
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            Message = "winget is not available on this system. Install ffmpeg manually or install App Installer from Microsoft Store."
        }
    }

    $arguments = "install Gyan.FFmpeg --accept-package-agreements --accept-source-agreements"
    $proc = Start-Process -FilePath $winget.Source -ArgumentList $arguments -PassThru -Wait -NoNewWindow

    $exitCode = 0
    try {
        $exitCode = [int]$proc.ExitCode
    }
    catch {
        $exitCode = -999
    }

    if ($exitCode -ne 0) {
        return [PSCustomObject]@{
            Success = $false
            Message = "winget install failed with exit code $exitCode."
        }
    }

    return [PSCustomObject]@{
        Success = $true
        Message = "ffmpeg installation completed."
    }
}

function Handle-StartupFFmpegCheck {
    param([string]$WhisperExePath)

    $ffmpegPath = Get-FFmpegPath -WhisperExePath $WhisperExePath
    if (-not [string]::IsNullOrWhiteSpace($ffmpegPath)) {
        return [PSCustomObject]@{
            Available = $true
            Path      = $ffmpegPath
            Message   = "ffmpeg detected at: $ffmpegPath"
        }
    }

    $prompt = @"
ffmpeg.exe was not found.

Whisper CLI on this build only reads wav/mp3/ogg/flac directly.
For m4a/aac/wma/opus/mp4/m4b, ffmpeg is required for conversion.

Do you want me to download/install ffmpeg Windows build now?
Yes = Install now (winget)
No = I will do it myself
"@

    $choice = [System.Windows.Forms.MessageBox]::Show(
        $prompt,
        "ffmpeg Not Found",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
        $installResult = Install-FFmpegWithWinget
        if (-not $installResult.Success) {
            [System.Windows.Forms.MessageBox]::Show($installResult.Message, "ffmpeg Install Failed", "OK", "Error") | Out-Null
            return [PSCustomObject]@{
                Available = $false
                Path      = $null
                Message   = $installResult.Message
            }
        }

        $ffmpegPath = Get-FFmpegPath -WhisperExePath $WhisperExePath
        if (-not [string]::IsNullOrWhiteSpace($ffmpegPath)) {
            [System.Windows.Forms.MessageBox]::Show("ffmpeg installed successfully.", "ffmpeg Installed", "OK", "Information") | Out-Null
            return [PSCustomObject]@{
                Available = $true
                Path      = $ffmpegPath
                Message   = "ffmpeg installed and detected at: $ffmpegPath"
            }
        }

        [System.Windows.Forms.MessageBox]::Show("ffmpeg install command completed, but ffmpeg was not detected yet. Restart the app after installation completes.", "ffmpeg Check", "OK", "Warning") | Out-Null
        return [PSCustomObject]@{
            Available = $false
            Path      = $null
            Message   = "ffmpeg install attempted, but ffmpeg is still not detected."
        }
    }

    return [PSCustomObject]@{
        Available = $false
        Path      = $null
        Message   = "ffmpeg not found. User chose to install manually."
    }
}

function Ensure-WhisperReadableAudio {
    param(
        [string]$AudioPath,
        [string]$WhisperExePath,
        [string]$TempDirectory
    )

    $ext = [System.IO.Path]::GetExtension($AudioPath).ToLowerInvariant()
    if ($script:NativeWhisperAudioExtensions -contains $ext) {
        return [PSCustomObject]@{
            InputPath   = $AudioPath
            Converted   = $false
            TempPath    = $null
            Message     = $null
            CanContinue = $true
        }
    }

    if (-not ($script:ConvertibleAudioExtensions -contains $ext)) {
        return [PSCustomObject]@{
            InputPath   = $AudioPath
            Converted   = $false
            TempPath    = $null
            Message     = "Unsupported file extension '$ext'. Supported native formats are: wav, mp3, ogg, flac."
            CanContinue = $false
        }
    }

    $ffmpegPath = Get-FFmpegPath -WhisperExePath $WhisperExePath
    if ([string]::IsNullOrWhiteSpace($ffmpegPath)) {
        return [PSCustomObject]@{
            InputPath   = $AudioPath
            Converted   = $false
            TempPath    = $null
            Message     = "ffmpeg.exe is required to convert '$ext' files. Install ffmpeg or use wav/mp3/ogg/flac input."
            CanContinue = $false
        }
    }

    if (-not (Test-Path -LiteralPath $TempDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $TempDirectory -Force | Out-Null
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($AudioPath)
    $wavPath = Join-Path $TempDirectory ("{0}_{1}.wav" -f $baseName, ([guid]::NewGuid().ToString("N")))
    $ffmpegArgs = @("-y", "-v", "error", "-i", $AudioPath, "-ar", "16000", "-ac", "1", $wavPath)
    $ffmpegArgString = ConvertTo-ProcessArgumentString -Arguments $ffmpegArgs

    $proc = Start-Process -FilePath $ffmpegPath -ArgumentList $ffmpegArgString -PassThru -Wait -NoNewWindow
    $exitCode = 0
    try {
        $exitCode = [int]$proc.ExitCode
    }
    catch {
        $exitCode = -999
    }

    if ($exitCode -ne 0 -or -not (Test-Path -LiteralPath $wavPath -PathType Leaf)) {
        if (Test-Path -LiteralPath $wavPath -PathType Leaf) {
            Remove-Item -LiteralPath $wavPath -Force -ErrorAction SilentlyContinue
        }

        return [PSCustomObject]@{
            InputPath   = $AudioPath
            Converted   = $false
            TempPath    = $null
            Message     = "ffmpeg conversion failed for '$AudioPath'."
            CanContinue = $false
        }
    }

    return [PSCustomObject]@{
        InputPath   = $wavPath
        Converted   = $true
        TempPath    = $wavPath
        Message     = ("Converted '{0}' to temporary WAV." -f $AudioPath)
        CanContinue = $true
    }
}

function Is-SupportedAudioFile {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }
    $ext = [System.IO.Path]::GetExtension($Path)
    if ([string]::IsNullOrWhiteSpace($ext)) {
        return $false
    }
    return $script:SupportedAudioExtensions -contains $ext.ToLowerInvariant()
}

function Get-SortedSelectedFiles {
    return @($script:SelectedAudioFiles | Sort-Object)
}

function Refresh-AudioList {
    param(
        [System.Windows.Forms.ListBox]$ListBox,
        [System.Windows.Forms.Label]$CountLabel
    )

    $files = Get-SortedSelectedFiles
    $ListBox.Items.Clear()
    foreach ($file in $files) {
        [void]$ListBox.Items.Add($file)
    }
    $CountLabel.Text = "Selected audio files: $($files.Count)"
}

function Add-AudioPaths {
    param([string[]]$Paths)

    $added = 0
    $skipped = 0

    foreach ($path in $Paths) {
        if ([string]::IsNullOrWhiteSpace($path)) {
            continue
        }

        try {
            $item = Get-Item -LiteralPath $path -ErrorAction Stop
            if ($item.PSIsContainer) {
                $skipped++
                continue
            }
            $fullPath = $item.FullName
        }
        catch {
            $skipped++
            continue
        }

        if (-not (Is-SupportedAudioFile -Path $fullPath)) {
            $skipped++
            continue
        }

        if ($script:SelectedAudioFiles.Add($fullPath)) {
            $added++
        }
    }

    return [PSCustomObject]@{
        Added   = $added
        Skipped = $skipped
    }
}

function Get-AudioFilesFromFolder {
    param(
        [string]$Folder,
        [bool]$Recursive = $false
    )

    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
        return @()
    }

    $items = Get-ChildItem -LiteralPath $Folder -File -Recurse:$Recursive -ErrorAction SilentlyContinue
    $files = New-Object System.Collections.Generic.List[string]
    foreach ($item in $items) {
        if (Is-SupportedAudioFile -Path $item.FullName) {
            $files.Add($item.FullName)
        }
    }

    return $files.ToArray()
}

function Set-UiEnabled {
    param(
        [bool]$Enabled,
        [System.Windows.Forms.Control[]]$Controls
    )

    foreach ($control in $Controls) {
        $control.Enabled = $Enabled
    }
}

function Get-DefaultWhisperExe {
    param([string]$BasePath)

    $candidates = @(
        "whisper-cli.exe",
        "whisper.exe"
    )

    foreach ($candidate in $candidates) {
        $full = Join-Path $BasePath $candidate
        if (Test-Path -LiteralPath $full) {
            return $full
        }
    }

    return (Join-Path $BasePath "whisper-bin-x64\Release\whisper-cli.exe")
}

function Get-OutputTarget {
    param(
        [string]$OutputFolder,
        [string]$BaseName,
        [System.Collections.Generic.HashSet[string]]$ReservedTxtPaths
    )

    $stem = $BaseName
    $index = 2

    while ($true) {
        $candidateTxt = Join-Path $OutputFolder "$stem.txt"
        if ($ReservedTxtPaths.Add($candidateTxt)) {
            return [PSCustomObject]@{
                OutputBase = [System.IO.Path]::Combine($OutputFolder, $stem)
                OutputTxt  = $candidateTxt
            }
        }

        $stem = "{0}_{1}" -f $BaseName, $index
        $index++
    }
}

function Invoke-WhisperExecution {
    param(
        [string]$ExePath,
        [string]$ArgumentString,
        [string]$StdOutPath,
        [string]$StdErrPath,
        [string]$StatusText,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $proc = Start-Process -FilePath $ExePath -ArgumentList $ArgumentString -PassThru -NoNewWindow -RedirectStandardOutput $StdOutPath -RedirectStandardError $StdErrPath

    while (-not $proc.HasExited) {
        [void]($StatusLabel.Text = $StatusText)
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 150
    }

    $proc.WaitForExit()

    $stdout = ""
    $stderr = ""
    if (Test-Path -LiteralPath $StdOutPath) {
        $stdout = Get-Content -LiteralPath $StdOutPath -Raw
    }
    if (Test-Path -LiteralPath $StdErrPath) {
        $stderr = Get-Content -LiteralPath $StdErrPath -Raw
    }

    $exitCode = 0
    try {
        $exitCode = [int]$proc.ExitCode
    }
    catch {
        $exitCode = -999
    }

    return [PSCustomObject]@{
        ExitCode = $exitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Whisper Windows Transcriber"
$form.Size = New-Object System.Drawing.Size(1040, 780)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false

$basePath = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { (Get-Location).Path }
$defaultWhisperExe = Get-DefaultWhisperExe -BasePath $basePath
$defaultOutputFolder = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "WhisperTranscripts"
$startupFFmpeg = Handle-StartupFFmpegCheck -WhisperExePath $defaultWhisperExe

$labelExe = New-Label -Text "Whisper EXE:" -X 20 -Y 24 -Width 130
$txtExe = New-TextBox -X 150 -Y 20 -Width 740
$txtExe.Text = $defaultWhisperExe
$btnExe = New-Object System.Windows.Forms.Button
$btnExe.Text = "Browse..."
$btnExe.Location = New-Object System.Drawing.Point(900, 18)
$btnExe.Size = New-Object System.Drawing.Size(110, 28)

$labelModel = New-Label -Text "Model (.bin):" -X 20 -Y 64 -Width 130
$txtModel = New-TextBox -X 150 -Y 60 -Width 740
$btnModel = New-Object System.Windows.Forms.Button
$btnModel.Text = "Browse..."
$btnModel.Location = New-Object System.Drawing.Point(900, 58)
$btnModel.Size = New-Object System.Drawing.Size(110, 28)

$labelAudio = New-Label -Text "Audio Sources:" -X 20 -Y 104 -Width 130
$btnAddFiles = New-Object System.Windows.Forms.Button
$btnAddFiles.Text = "Add Files..."
$btnAddFiles.Location = New-Object System.Drawing.Point(150, 100)
$btnAddFiles.Size = New-Object System.Drawing.Size(120, 28)

$btnAddFolder = New-Object System.Windows.Forms.Button
$btnAddFolder.Text = "Add Folder..."
$btnAddFolder.Location = New-Object System.Drawing.Point(280, 100)
$btnAddFolder.Size = New-Object System.Drawing.Size(120, 28)

$btnClearFiles = New-Object System.Windows.Forms.Button
$btnClearFiles.Text = "Clear List"
$btnClearFiles.Location = New-Object System.Drawing.Point(410, 100)
$btnClearFiles.Size = New-Object System.Drawing.Size(100, 28)

$chkRecursive = New-Object System.Windows.Forms.CheckBox
$chkRecursive.Text = "Include subfolders when adding a folder"
$chkRecursive.Location = New-Object System.Drawing.Point(525, 103)
$chkRecursive.Size = New-Object System.Drawing.Size(300, 24)
$chkRecursive.Checked = $false

$labelCount = New-Object System.Windows.Forms.Label
$labelCount.Location = New-Object System.Drawing.Point(840, 104)
$labelCount.Size = New-Object System.Drawing.Size(170, 24)
$labelCount.TextAlign = "MiddleRight"
$labelCount.Text = "Selected audio files: 0"

$labelAudioFormats = New-Object System.Windows.Forms.Label
$labelAudioFormats.Location = New-Object System.Drawing.Point(150, 132)
$labelAudioFormats.Size = New-Object System.Drawing.Size(860, 18)
$labelAudioFormats.Text = "Whisper native formats: wav, mp3, ogg, flac | Converted with ffmpeg: m4a, aac, wma, opus, mp4, m4b"
if (-not $startupFFmpeg.Available) {
    $labelAudioFormats.ForeColor = [System.Drawing.Color]::DarkRed
}

$listAudio = New-Object System.Windows.Forms.ListBox
$listAudio.Location = New-Object System.Drawing.Point(150, 152)
$listAudio.Size = New-Object System.Drawing.Size(860, 100)
$listAudio.HorizontalScrollbar = $true

$labelOutputFolder = New-Label -Text "Output Folder:" -X 20 -Y 264 -Width 130
$txtOutputFolder = New-TextBox -X 150 -Y 260 -Width 740
$txtOutputFolder.Text = $defaultOutputFolder
$btnOutputFolder = New-Object System.Windows.Forms.Button
$btnOutputFolder.Text = "Browse..."
$btnOutputFolder.Location = New-Object System.Drawing.Point(900, 258)
$btnOutputFolder.Size = New-Object System.Drawing.Size(110, 28)

$labelPrompt = New-Label -Text "Initial Prompt:" -X 20 -Y 304 -Width 130
$txtPrompt = New-Object System.Windows.Forms.TextBox
$txtPrompt.Location = New-Object System.Drawing.Point(150, 300)
$txtPrompt.Size = New-Object System.Drawing.Size(860, 70)
$txtPrompt.Multiline = $true
$txtPrompt.ScrollBars = "Vertical"
$txtPrompt.Text = "Transcribe this audio faithfully. Segment by speaker turns and format as Speaker 1:, Speaker 2:, etc. Preserve meaning, key pauses, hesitations, and sentence boundaries."

$labelLanguage = New-Label -Text "Language:" -X 20 -Y 380 -Width 130
$cmbLanguage = New-Object System.Windows.Forms.ComboBox
$cmbLanguage.Location = New-Object System.Drawing.Point(150, 376)
$cmbLanguage.Size = New-Object System.Drawing.Size(280, 26)
$cmbLanguage.DropDownStyle = "DropDownList"
$cmbLanguage.DisplayMember = "Label"
$cmbLanguage.ValueMember = "Code"
[void]$cmbLanguage.Items.AddRange((Get-WhisperLanguageOptions))
$cmbLanguage.SelectedIndex = 0

$chkTranslate = New-Object System.Windows.Forms.CheckBox
$chkTranslate.Text = "Translate to English (-tr)"
$chkTranslate.Location = New-Object System.Drawing.Point(450, 378)
$chkTranslate.Size = New-Object System.Drawing.Size(260, 24)
$chkTranslate.Checked = $false

$chkDiarize = New-Object System.Windows.Forms.CheckBox
$chkDiarize.Text = "Enable speaker turn diarization (-tdrz, if your whisper build supports it)"
$chkDiarize.Location = New-Object System.Drawing.Point(150, 410)
$chkDiarize.Size = New-Object System.Drawing.Size(520, 26)
$chkDiarize.Checked = $true

$chkNoTimestamps = New-Object System.Windows.Forms.CheckBox
$chkNoTimestamps.Text = "Suppress timestamps in output (-nt)"
$chkNoTimestamps.Location = New-Object System.Drawing.Point(700, 410)
$chkNoTimestamps.Size = New-Object System.Drawing.Size(310, 26)
$chkNoTimestamps.Checked = $false

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Start Transcription"
$btnRun.Location = New-Object System.Drawing.Point(150, 444)
$btnRun.Size = New-Object System.Drawing.Size(220, 38)

$status = New-Object System.Windows.Forms.Label
$status.Text = "Status: Ready"
$status.Location = New-Object System.Drawing.Point(390, 454)
$status.Size = New-Object System.Drawing.Size(620, 24)

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(20, 490)
$logBox.Size = New-Object System.Drawing.Size(990, 240)
$logBox.ReadOnly = $true
$logBox.ScrollBars = "Vertical"
$logBox.WordWrap = $false

$openExe = New-Object System.Windows.Forms.OpenFileDialog
$openExe.Filter = "Executable (*.exe)|*.exe|All files (*.*)|*.*"
$openExe.Title = "Select whisper executable (prefer whisper-cli.exe)"

$openModel = New-Object System.Windows.Forms.OpenFileDialog
$openModel.Filter = "Whisper model (*.bin)|*.bin|All files (*.*)|*.*"
$openModel.Title = "Select whisper model"

$openAudio = New-Object System.Windows.Forms.OpenFileDialog
$openAudio.Filter = "Audio files (*.wav;*.mp3;*.m4a;*.flac;*.ogg;*.aac;*.wma;*.opus;*.mp4;*.m4b)|*.wav;*.mp3;*.m4a;*.flac;*.ogg;*.aac;*.wma;*.opus;*.mp4;*.m4b|All files (*.*)|*.*"
$openAudio.Title = "Select one or more audio files"
$openAudio.Multiselect = $true

$folderPicker = New-Object System.Windows.Forms.FolderBrowserDialog
$folderPicker.ShowNewFolderButton = $true

$controlsToDisableWhenRunning = @(
    $txtExe, $txtModel, $btnExe, $btnModel,
    $btnAddFiles, $btnAddFolder, $btnClearFiles, $chkRecursive,
    $txtOutputFolder, $btnOutputFolder,
    $txtPrompt, $cmbLanguage, $chkTranslate, $chkDiarize, $chkNoTimestamps,
    $btnRun
)

$btnExe.Add_Click({
    if ($openExe.ShowDialog() -eq "OK") {
        $txtExe.Text = $openExe.FileName
    }
})

$btnModel.Add_Click({
    if ($openModel.ShowDialog() -eq "OK") {
        $txtModel.Text = $openModel.FileName
    }
})

$btnAddFiles.Add_Click({
    if ($openAudio.ShowDialog() -eq "OK") {
        $result = Add-AudioPaths -Paths $openAudio.FileNames
        Refresh-AudioList -ListBox $listAudio -CountLabel $labelCount
        Add-Log -LogTextBox $logBox -Message ("Added files: {0}, skipped: {1}" -f $result.Added, $result.Skipped)
        foreach ($file in (Get-SortedSelectedFiles)) {
            Add-Log -LogTextBox $logBox -Message ("Selected: {0}" -f $file)
        }
    }
})

$btnAddFolder.Add_Click({
    $folderPicker.Description = "Select a folder containing audio files"
    if ($folderPicker.ShowDialog() -eq "OK") {
        $folderFiles = Get-AudioFilesFromFolder -Folder $folderPicker.SelectedPath -Recursive:$chkRecursive.Checked
        $result = Add-AudioPaths -Paths $folderFiles
        Refresh-AudioList -ListBox $listAudio -CountLabel $labelCount
        Add-Log -LogTextBox $logBox -Message ("Folder added: {0}" -f $folderPicker.SelectedPath)
        Add-Log -LogTextBox $logBox -Message ("Added files: {0}, skipped: {1}" -f $result.Added, $result.Skipped)
        foreach ($file in (Get-SortedSelectedFiles)) {
            Add-Log -LogTextBox $logBox -Message ("Selected: {0}" -f $file)
        }

        if ($result.Added -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No supported audio files found in the selected folder.", "No Audio Found", "OK", "Information") | Out-Null
        }
    }
})

$btnClearFiles.Add_Click({
    $script:SelectedAudioFiles.Clear()
    Refresh-AudioList -ListBox $listAudio -CountLabel $labelCount
    Add-Log -LogTextBox $logBox -Message "Cleared selected audio list."
})

$btnOutputFolder.Add_Click({
    $folderPicker.Description = "Select output folder for transcript .txt files"
    if (-not [string]::IsNullOrWhiteSpace($txtOutputFolder.Text) -and (Test-Path -LiteralPath $txtOutputFolder.Text -PathType Container)) {
        $folderPicker.SelectedPath = $txtOutputFolder.Text
    }
    if ($folderPicker.ShowDialog() -eq "OK") {
        $txtOutputFolder.Text = $folderPicker.SelectedPath
    }
})

$btnRun.Add_Click({
    $exePath = $txtExe.Text.Trim()
    $modelPath = $txtModel.Text.Trim()
    $outputFolder = $txtOutputFolder.Text.Trim()
    $prompt = $txtPrompt.Text.Trim()
    $languageCode = "auto"
    $languageLabel = "Auto Detect (auto)"
    if ($cmbLanguage.SelectedItem -and $cmbLanguage.SelectedItem.PSObject.Properties["Code"]) {
        $candidateCode = [string]$cmbLanguage.SelectedItem.Code
        if (-not [string]::IsNullOrWhiteSpace($candidateCode)) {
            $languageCode = $candidateCode
        }
        if ($cmbLanguage.SelectedItem.PSObject.Properties["Label"]) {
            $languageLabel = [string]$cmbLanguage.SelectedItem.Label
        }
    }
    $translateEnabled = $chkTranslate.Checked
    $audioFiles = @(Get-SortedSelectedFiles)

    if ([string]::IsNullOrWhiteSpace($exePath) -or -not (Test-Path -LiteralPath $exePath)) {
        [System.Windows.Forms.MessageBox]::Show("Select a valid whisper executable path.", "Missing EXE", "OK", "Warning") | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($modelPath) -or -not (Test-Path -LiteralPath $modelPath)) {
        [System.Windows.Forms.MessageBox]::Show("Select a valid model .bin file.", "Missing Model", "OK", "Warning") | Out-Null
        return
    }
    if ($audioFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Add at least one audio file or folder.", "Missing Audio", "OK", "Warning") | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($outputFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Choose an output folder for transcript files.", "Missing Output Folder", "OK", "Warning") | Out-Null
        return
    }

    try {
        if (-not (Test-Path -LiteralPath $outputFolder -PathType Container)) {
            New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Could not create output folder: $($_.Exception.Message)", "Output Folder Error", "OK", "Error") | Out-Null
        return
    }

    Set-UiEnabled -Enabled $false -Controls $controlsToDisableWhenRunning
    $status.Text = "Status: Running..."
    Add-Log -LogTextBox $logBox -Message ("Starting transcription for {0} file(s)..." -f $audioFiles.Count)
    Add-Log -LogTextBox $logBox -Message ("Language: {0} | Code: {1} | Translate: {2}" -f $languageLabel, $languageCode, $(if ($translateEnabled) { "On" } else { "Off" }))

    $reservedTxtPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $successCount = 0
    $failureCount = 0

    try {
        $totalFiles = $audioFiles.Count
        $index = 0

        foreach ($audioPath in $audioFiles) {
            $index++

            if (-not (Test-Path -LiteralPath $audioPath -PathType Leaf)) {
                $failureCount++
                Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] Failed: invalid audio path: $audioPath")
                continue
            }

            $tempAudioDir = Join-Path $env:TEMP "whisper_ui_audio"
            $audioPrep = $null
            try {
                $audioPrep = Ensure-WhisperReadableAudio -AudioPath $audioPath -WhisperExePath $exePath -TempDirectory $tempAudioDir
            }
            catch {
                $failureCount++
                Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] Failed while preparing audio: {0}" -f $_.Exception.Message)
                continue
            }

            if (-not $audioPrep.CanContinue) {
                $failureCount++
                Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] Failed: {0}" -f $audioPrep.Message)
                continue
            }

            if ($audioPrep.Converted) {
                Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] {0}" -f $audioPrep.Message)
            }

            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($audioPath)
            $outputTarget = Get-OutputTarget -OutputFolder $outputFolder -BaseName $baseName -ReservedTxtPaths $reservedTxtPaths

            $outputBase = $outputTarget.OutputBase
            $outputTxt = $outputTarget.OutputTxt

            $arguments = @(
                "-m", $modelPath,
                "-f", $audioPrep.InputPath,
                "-otxt",
                "-of", $outputBase
            )
            if (-not [string]::IsNullOrWhiteSpace($languageCode)) {
                $arguments += @("-l", $languageCode)
            }
            if ($translateEnabled) {
                $arguments += "-tr"
            }
            if (-not [string]::IsNullOrWhiteSpace($prompt)) {
                $arguments += @("--prompt", $prompt)
            }
            if ($chkNoTimestamps.Checked) {
                $arguments += "-nt"
            }
            if ($chkDiarize.Checked) {
                $arguments += "-tdrz"
            }

            $status.Text = "Status: Running ($index/$totalFiles)"
            Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] Transcribing: $audioPath")
            Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] Output: $outputTxt")

            $tempOut = Join-Path $env:TEMP ("whisper_stdout_{0}.log" -f ([guid]::NewGuid().ToString("N")))
            $tempErr = Join-Path $env:TEMP ("whisper_stderr_{0}.log" -f ([guid]::NewGuid().ToString("N")))

            try {
                Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] Command: `"{0}`" {1}" -f $exePath, (Format-ArgsForLog -Arguments $arguments))
                $argumentString = ConvertTo-ProcessArgumentString -Arguments $arguments
                $result = Invoke-WhisperExecution -ExePath $exePath -ArgumentString $argumentString -StdOutPath $tempOut -StdErrPath $tempErr -StatusText ("Status: Running ($index/$totalFiles)") -StatusLabel $status

                if (-not [string]::IsNullOrWhiteSpace($result.StdOut)) {
                    Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] whisper stdout:")
                    Add-Log -LogTextBox $logBox -Message $result.StdOut.Trim()
                }
                if (-not [string]::IsNullOrWhiteSpace($result.StdErr)) {
                    Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] whisper stderr:")
                    Add-Log -LogTextBox $logBox -Message $result.StdErr.Trim()
                }

                if ($result.ExitCode -ne 0 -and $chkDiarize.Checked -and $arguments -contains "-tdrz" -and $result.StdErr -match "(?i)((unknown|unrecognized|unsupported|invalid)\s+(option|argument).*(-tdrz|--tinydiarize)|(requires a tdrz model))") {
                    Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] Retrying without -tdrz (diarization flag may be unsupported in this build).")
                    $arguments = @($arguments | Where-Object { $_ -ne "-tdrz" })
                    Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] Retry command: `"{0}`" {1}" -f $exePath, (Format-ArgsForLog -Arguments $arguments))
                    $argumentString = ConvertTo-ProcessArgumentString -Arguments $arguments
                    $result = Invoke-WhisperExecution -ExePath $exePath -ArgumentString $argumentString -StdOutPath $tempOut -StdErrPath $tempErr -StatusText ("Status: Running ($index/$totalFiles) retry") -StatusLabel $status

                    if (-not [string]::IsNullOrWhiteSpace($result.StdOut)) {
                        Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] retry stdout:")
                        Add-Log -LogTextBox $logBox -Message $result.StdOut.Trim()
                    }
                    if (-not [string]::IsNullOrWhiteSpace($result.StdErr)) {
                        Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] retry stderr:")
                        Add-Log -LogTextBox $logBox -Message $result.StdErr.Trim()
                    }
                }

                if ($result.ExitCode -ne 0) {
                    throw "whisper executable exited with code $($result.ExitCode)."
                }
                if (-not (Test-Path -LiteralPath $outputTxt)) {
                    throw "Expected transcript file not found: $outputTxt"
                }

                $successCount++
                Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] Completed")
            }
            catch {
                $failureCount++
                Add-Log -LogTextBox $logBox -Message ("[$index/$totalFiles] Failed: {0}" -f $_.Exception.Message)
            }
            finally {
                if (Test-Path -LiteralPath $tempOut) {
                    Remove-Item -LiteralPath $tempOut -Force -ErrorAction SilentlyContinue
                }
                if (Test-Path -LiteralPath $tempErr) {
                    Remove-Item -LiteralPath $tempErr -Force -ErrorAction SilentlyContinue
                }
                if ($audioPrep -and $audioPrep.Converted -and -not [string]::IsNullOrWhiteSpace($audioPrep.TempPath) -and (Test-Path -LiteralPath $audioPrep.TempPath -PathType Leaf)) {
                    Remove-Item -LiteralPath $audioPrep.TempPath -Force -ErrorAction SilentlyContinue
                }
            }
        }

        if ($failureCount -eq 0) {
            $status.Text = "Status: Completed"
            Add-Log -LogTextBox $logBox -Message ("Finished successfully. Files completed: {0}" -f $successCount)
            [System.Windows.Forms.MessageBox]::Show("Transcription complete.`r`n`r`nFiles completed: $successCount`r`nOutput folder: $outputFolder", "Success", "OK", "Information") | Out-Null
        }
        else {
            $status.Text = "Status: Completed with errors"
            Add-Log -LogTextBox $logBox -Message ("Finished with errors. Success: {0}, Failed: {1}" -f $successCount, $failureCount)
            [System.Windows.Forms.MessageBox]::Show("Transcription finished with errors.`r`n`r`nSuccess: $successCount`r`nFailed: $failureCount`r`nOutput folder: $outputFolder`r`n`r`nCheck the log for details.", "Completed with Errors", "OK", "Warning") | Out-Null
        }
    }
    finally {
        Set-UiEnabled -Enabled $true -Controls $controlsToDisableWhenRunning
    }
})

$form.Controls.AddRange(@(
    $labelExe, $txtExe, $btnExe,
    $labelModel, $txtModel, $btnModel,
    $labelAudio, $btnAddFiles, $btnAddFolder, $btnClearFiles, $chkRecursive, $labelCount,
    $labelAudioFormats,
    $listAudio,
    $labelOutputFolder, $txtOutputFolder, $btnOutputFolder,
    $labelPrompt, $txtPrompt, $labelLanguage, $cmbLanguage, $chkTranslate,
    $chkDiarize, $chkNoTimestamps,
    $btnRun, $status, $logBox
))

Add-Log -LogTextBox $logBox -Message "Choose whisper-cli.exe, your large-v3 q5/q8 model, and add files or folders."
Add-Log -LogTextBox $logBox -Message "Output .txt names are auto-generated from each audio filename."
Add-Log -LogTextBox $logBox -Message "For m4a/aac/wma/opus/mp4/m4b input, ffmpeg is used to convert to temporary WAV."
Add-Log -LogTextBox $logBox -Message "Choose transcription language from dropdown, or keep Auto Detect."
Add-Log -LogTextBox $logBox -Message ("Startup ffmpeg check: {0}" -f $startupFFmpeg.Message)

[void]$form.ShowDialog()
