<#
.SYNOPSIS
    安裝影片逐字稿工具所需的相依套件。

.DESCRIPTION
    檢查並安裝：yt-dlp（抓字幕／音訊）、FFmpeg（音訊轉檔）、
    Python 3.12、faster-whisper（本機語音辨識）、opencc（簡轉繁）。
    已安裝的項目會自動略過，重複執行是安全的。

.EXAMPLE
    .\setup.ps1

.EXAMPLE
    # 順便先把 Whisper 模型下載好，之後第一次辨識就不用等
    .\setup.ps1 -PrefetchModel large-v3-turbo
#>
[CmdletBinding()]
param(
    [string] $PrefetchModel = '',
    [switch] $Upgrade
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

function Write-Step { param([string]$Text) Write-Host ("`n== $Text") -ForegroundColor Cyan }
function Write-Ok   { param([string]$Text) Write-Host ("   [完成] $Text") -ForegroundColor Green }
function Write-Info { param([string]$Text) Write-Host ("   $Text") -ForegroundColor DarkGray }
function Write-Warn { param([string]$Text) Write-Host ("   [注意] $Text") -ForegroundColor Yellow }

function Update-PathFromRegistry {
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user    = [Environment]::GetEnvironmentVariable('Path', 'User')
    if ($machine -or $user) { $env:Path = "$machine;$user" }
}

function Get-ToolPath {
    param([string]$Name, [string[]]$Hints = @())
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    foreach ($h in $Hints) {
        if (-not (Test-Path $h)) { continue }
        $hit = Get-ChildItem -Path $h -Filter $Name -Recurse -ErrorAction SilentlyContinue |
               Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    return $null
}

function Install-WingetPackage {
    param([string]$Id, [string]$Exe, [string[]]$Hints, [string]$Label, [string[]]$Extra = @())

    Update-PathFromRegistry
    $found = Get-ToolPath -Name $Exe -Hints $Hints
    if ($found -and -not $Upgrade) {
        Write-Ok "$Label 已安裝：$found"
        return $found
    }

    Write-Info "以 winget 安裝 $Label…"
    $wingetArgs = @('install', '--id', $Id, '-e',
                    '--accept-package-agreements', '--accept-source-agreements',
                    '--disable-interactivity') + $Extra
    & winget @wingetArgs | Out-Null

    Update-PathFromRegistry
    $found = Get-ToolPath -Name $Exe -Hints $Hints
    if ($found) { Write-Ok "$Label：$found" } else { Write-Warn "$Label 安裝後仍找不到，請手動確認。" }
    return $found
}

# ---------------------------------------------------------------------------

Write-Host '影片逐字稿工具 — 環境安裝' -ForegroundColor White

Write-Step '檢查 winget'
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    throw '找不到 winget。請先從 Microsoft Store 更新「應用程式安裝程式」後再執行。'
}
Write-Ok ('winget ' + (& winget --version))

$wingetPkgs = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'
$pyRoot     = Join-Path $env:LOCALAPPDATA 'Programs\Python'

Write-Step '安裝 yt-dlp（抓取字幕與音訊）'
$ytdlp = Install-WingetPackage -Id 'yt-dlp.yt-dlp' -Exe 'yt-dlp.exe' `
            -Hints @($wingetPkgs) -Label 'yt-dlp'

Write-Step '安裝 FFmpeg（音訊轉檔）'
$ffmpeg = Install-WingetPackage -Id 'Gyan.FFmpeg' -Exe 'ffmpeg.exe' `
            -Hints @($wingetPkgs) -Label 'FFmpeg'

Write-Step '安裝 Python 3.12（供 Whisper 使用）'
$python = Install-WingetPackage -Id 'Python.Python.3.12' -Exe 'python.exe' `
            -Hints @($pyRoot) -Label 'Python' -Extra @('--scope', 'user')

if ($python) {
    Write-Step '安裝 faster-whisper 與 opencc'
    & $python -m pip install --upgrade --quiet pip
    & $python -m pip install --upgrade --quiet faster-whisper opencc-python-reimplemented
    if ($LASTEXITCODE -eq 0) {
        $ver = & $python -c "import faster_whisper as fw; print(fw.__version__)"
        Write-Ok "faster-whisper $ver"
    } else {
        Write-Warn 'pip 安裝失敗，語音辨識功能將無法使用（字幕抓取仍可正常運作）。'
    }

    if ($PrefetchModel) {
        Write-Step "預先下載 Whisper 模型：$PrefetchModel"
        Write-Info '首次下載約需數分鐘，模型會存到使用者目錄的 HuggingFace 快取。'
        & $python -c "from faster_whisper import WhisperModel; WhisperModel('$PrefetchModel', device='cpu', compute_type='int8'); print('ok')"
        if ($LASTEXITCODE -eq 0) { Write-Ok "模型 $PrefetchModel 已就緒" } else { Write-Warn '模型下載失敗，第一次辨識時會自動重試。' }
    }
}

Write-Step '安裝結果'
Write-Host ('   yt-dlp ：{0}' -f $(if ($ytdlp)  { $ytdlp }  else { '未安裝' }))
Write-Host ('   ffmpeg ：{0}' -f $(if ($ffmpeg) { $ffmpeg } else { '未安裝' }))
Write-Host ('   python ：{0}' -f $(if ($python) { $python } else { '未安裝' }))

Write-Host ''
Write-Host '完成。用法範例：' -ForegroundColor Cyan
Write-Host '   .\Get-VideoTranscript.ps1 -Url "https://www.youtube.com/watch?v=XXXXXXXXXXX"' -ForegroundColor White
Write-Host ''
