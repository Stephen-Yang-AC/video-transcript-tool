<#
.SYNOPSIS
    將線上影片轉成逐字稿（YouTube、教育大市集，以及 yt-dlp 支援的其他網站）。

.DESCRIPTION
    優先抓取影片既有字幕（人工字幕 > 自動字幕）；若完全沒有字幕，
    才下載音訊改用本機 Whisper 語音辨識。輸出五種檔案：
      1. _逐字稿_整理版.txt   補標點、分段、去口頭禪，可直接引用
      2. _逐字稿_原始.txt     忠實呈現字幕原文，不做改寫
      3. _逐字稿_時間碼.txt   每段前加 [hh:mm:ss]，便於回頭核對原片
      4. .srt                 標準字幕檔
      5. _逐字稿.docx         標楷體 + Times New Roman、含頁碼，整理版與原始版各一節

.EXAMPLE
    .\Get-VideoTranscript.ps1 -Url "https://www.youtube.com/watch?v=XXXXXXXXXXX"

.EXAMPLE
    .\Get-VideoTranscript.ps1 -Url "https://youtu.be/XXXX" -ForceWhisper -Model large-v3
#>
[CmdletBinding()]
param(
    # ValueFromRemainingArguments 讓多個網址能以空白分隔傳入。
    # 這是啟動器 逐字稿.cmd 能支援多支影片的關鍵：powershell -File 只會把
    # 「-Url "a","b"」當成一個字串，得靠位置參數逐個接收才行。
    [Parameter(Mandatory = $true, Position = 0, ValueFromRemainingArguments = $true)]
    [string[]] $Url,

    [string] $OutDir,

    # 字幕語言優先序，支援 yt-dlp 的萬用字元寫法
    [string] $SubLangs = 'zh.*,en.*',

    # 略過字幕、直接用語音辨識
    [switch] $ForceWhisper,

    # Whisper 模型：tiny / base / small / medium / large-v3-turbo / large-v3
    [string] $Model = 'large-v3-turbo',

    # 語音語言，留空為自動偵測；中文建議填 zh
    [string] $WhisperLang = 'zh',

    [ValidateSet('int8', 'int8_float32', 'float32', 'float16')]
    [string] $ComputeType = 'int8',

    # 會員限定或需登入的影片：指定瀏覽器借用 cookie，如 edge / chrome
    [string] $CookiesFromBrowser = '',

    [switch] $NoDocx,
    [switch] $KeepTemp
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

# ValueFromRemainingArguments 會把無法辨識的參數一併收進 $Url，
# 這裡先攔下來，免得打錯的旗標被當成網址、丟出看不懂的錯誤。
$badArgs = @($Url | Where-Object { $_ -like '-*' })
if ($badArgs.Count -gt 0) {
    Write-Host ('無法辨識的參數：{0}' -f ($badArgs -join '、')) -ForegroundColor Red
    Write-Host '請檢查拼寫。可用參數見 README 第四節，或執行：' -ForegroundColor Red
    Write-Host '   Get-Help .\Get-VideoTranscript.ps1 -Detailed' -ForegroundColor Red
    exit 1
}

$script:Root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $OutDir) { $OutDir = Join-Path $script:Root '輸出' }

$script:SourceLabel = ''
$script:PickedLang  = ''

# ---------------------------------------------------------------- 工具定位 --

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

function Initialize-Tools {
    # 剛安裝完相依套件時，目前這個 shell 的 PATH 可能還是舊的，先重讀一次
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user    = [Environment]::GetEnvironmentVariable('Path', 'User')
    if ($machine -or $user) { $env:Path = "$machine;$user" }

    $wingetPkgs = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'
    $pyRoot     = Join-Path $env:LOCALAPPDATA 'Programs\Python'

    $script:YtDlp  = Get-ToolPath -Name 'yt-dlp.exe' -Hints @($wingetPkgs)
    $script:FFmpeg = Get-ToolPath -Name 'ffmpeg.exe' -Hints @($wingetPkgs)
    $script:Python = Get-ToolPath -Name 'python.exe' -Hints @($pyRoot)

    if (-not $script:YtDlp) {
        throw '找不到 yt-dlp，請先執行 setup.ps1 安裝相依套件。'
    }
}

# ------------------------------------------------------------ 外部程式呼叫 --

function Invoke-Native {
    # 呼叫外部執行檔。yt-dlp 與 Whisper 都會把進度寫到 stderr，
    # 在 $ErrorActionPreference = 'Stop' 之下那會被當成終止錯誤，所以這裡暫時放寬。
    param([string]$Exe, [string[]]$Arguments, [switch]$Quiet)
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        if ($Quiet) { & $Exe @Arguments 2>&1 | Out-Null }
        else        { & $Exe @Arguments }
    }
    finally { $ErrorActionPreference = $prev }
}

# -------------------------------------------------------------- 來源解析 --

function Resolve-VideoSource {
    <#
      yt-dlp 認得上千個網站，但不認得教育大市集（market.cloud.edu.tw）。
      該站的播放器是讀頁面上的 hex_path 欄位：欄位若是 YouTube 網址就嵌 YouTube，
      否則組成 <StreamingURL>/medialib/<hex_path>/playlist.m3u8 走 HLS 串流。
      這裡照同樣規則把真正的影音位址解出來，再交給 yt-dlp。
    #>
    param([string]$InputUrl)

    $out = [pscustomobject]@{ Url = $InputUrl; Title = ''; Site = ''; Note = '' }
    if ($InputUrl -notmatch 'market\.cloud\.edu\.tw') { return $out }

    Write-Host '  - 偵測到教育大市集網址，解析影音來源…' -ForegroundColor DarkGray
    try {
        $resp = Invoke-WebRequest -Uri $InputUrl -UseBasicParsing -TimeoutSec 60
    }
    catch {
        throw ('無法讀取教育大市集頁面：{0}' -f $_.Exception.Message)
    }

    # PowerShell 5.1 判讀網頁編碼並不可靠，直接以 UTF-8 解位元組
    $html = $resp.Content
    if ($resp.RawContentStream) {
        try {
            $html = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
        } catch { }
    }

    $out.Site = '教育大市集'
    if ($html -match '(?s)<title>\s*(.*?)\s*</title>') { $out.Title = $Matches[1] }

    $hex = ''
    $stream = ''
    if ($html -match 'id="hex_path"[^>]*value="([^"]*)"')     { $hex    = $Matches[1].Trim() }
    if ($html -match 'id="StreamingURL"[^>]*value="([^"]*)"') { $stream = $Matches[1].Trim() }

    # hex_path 也可能直接放 YouTube 網址或影片 ID
    if ($hex -match '(?:youtube\.com/.*[?&]v=|youtu\.be/|youtube\.com/embed/)([A-Za-z0-9_-]{11})') {
        $out.Url  = 'https://www.youtube.com/watch?v=' + $Matches[1]
        $out.Note = '教育大市集（影片實際掛在 YouTube）'
        return $out
    }
    if ($hex -match '^[A-Za-z0-9_-]{11}$') {
        $out.Url  = 'https://www.youtube.com/watch?v=' + $hex
        $out.Note = '教育大市集（影片實際掛在 YouTube）'
        return $out
    }

    if (-not $hex -or -not $stream) {
        throw '在教育大市集頁面上找不到影音位址，該資源可能不是影片、或需要登入。'
    }

    $out.Url  = '{0}/medialib/{1}/playlist.m3u8' -f $stream.TrimEnd('/'), $hex.Trim('/')
    $out.Note = '教育大市集（HLS 串流，此平台不提供字幕，將以語音辨識產生逐字稿）'
    return $out
}

# ------------------------------------------------------------ 文字處理工具 --

function Test-CjkChar {
    param([char]$Char)
    $v = [int]$Char
    return (($v -ge 0x3000 -and $v -le 0x303F) -or   # CJK 標點
            ($v -ge 0x3400 -and $v -le 0x4DBF) -or   # 擴充 A
            ($v -ge 0x4E00 -and $v -le 0x9FFF) -or   # 基本漢字
            ($v -ge 0xFF00 -and $v -le 0xFFEF))      # 全形符號
}

function Join-Text {
    param([string]$Left, [string]$Right)
    if ([string]::IsNullOrEmpty($Left))  { return $Right }
    if ([string]::IsNullOrEmpty($Right)) { return $Left }
    $a = $Left[$Left.Length - 1]
    $b = $Right[0]
    # 中文之間直接相連，英數之間才補空格
    if ((Test-CjkChar $a) -or (Test-CjkChar $b)) { return $Left + $Right }
    return $Left + ' ' + $Right
}

function ConvertTo-Seconds {
    param([string]$Stamp)
    $s = $Stamp.Trim() -replace ',', '.'
    $sec = 0.0
    foreach ($p in ($s -split ':')) {
        $v = 0.0
        if (-not [double]::TryParse($p, [ref]$v)) { $v = 0.0 }   # 來源時間碼可能殘缺
        $sec = $sec * 60 + $v
    }
    return $sec
}

function Format-Timestamp {
    param([double]$Seconds, [string]$Style = 'clock')
    $ts = [TimeSpan]::FromSeconds($Seconds)
    # 注意：PowerShell 的 [int] 轉型是四捨五入而非捨去，
    # 直接用 [int]$ts.TotalHours 會把 58 分鐘變成 1 小時，必須用 Floor。
    $hours = [int][math]::Floor($ts.TotalHours)
    if ($Style -eq 'srt') {
        return '{0:00}:{1:00}:{2:00},{3:000}' -f $hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds
    }
    return '{0:00}:{1:00}:{2:00}' -f $hours, $ts.Minutes, $ts.Seconds
}

function Clear-CueMarkup {
    param([string]$Line)
    if ($null -eq $Line) { return '' }
    $t = $Line -replace '<[^>]*>', ''            # <00:00:01.234> 與 <c> 標記
    $t = $t -replace '&amp;',  '&'
    $t = $t -replace '&lt;',   '<'
    $t = $t -replace '&gt;',   '>'
    $t = $t -replace '&quot;', '"'
    $t = $t -replace '&#39;',  [char]39
    $t = $t -replace '&nbsp;', ' '
    $t = $t -replace '\s{2,}', ' '
    $t = $t.Trim()

    # 第二道防線：整行只由數字、冒號與箭頭構成者一律視為字幕檔標記，
    # 連同 SRT 誤貼進來的序號行一併濾掉。
    if ($t -match '^[\d:,\.\s]*-->[\d:,\.\s]*$') { return '' }
    if ($t -match '^\d{1,5}$') { return '' }

    return $t
}

function Clear-Filler {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $t = $Text -replace '[呃嗯]+', ''                                     # 語助詞
    $t = $t -replace '(那個|這個|就是|然後|所以|其實|對|好|欸)\1+', '$1'   # 立即重複
    $t = $t -replace '\s{2,}', ' '
    # 中文字幕常以空格代替逗號（例如「好 非常謝謝部長」），整理版直接補回逗號
    $t = $t -replace '(?<=[一-鿿])\s+(?=[一-鿿])', '，'
    return $t.Trim()
}

function Convert-ToFullWidthPunct {
    param([string]$Text)
    $t = $Text
    $t = $t -replace '(?<=[一-鿿]),',  '，'
    $t = $t -replace '(?<=[一-鿿]);',  '；'
    $t = $t -replace '(?<=[一-鿿]):',  '：'
    $t = $t -replace '(?<=[一-鿿])\?', '？'
    $t = $t -replace '(?<=[一-鿿])!',  '！'
    $t = $t -replace '，{2,}', '，'
    $t = $t -replace '。{2,}', '。'
    return $t
}

function Close-Sentence {
    param([string]$Text)
    $t = $Text.TrimEnd()
    $t = $t -replace '[，、；：\s]+$', ''
    if ($t -eq '') { return '' }
    $last = $t[$t.Length - 1]
    if ('。！？」』）.!?'.IndexOf($last) -lt 0) { $t += '。' }
    return (Convert-ToFullWidthPunct $t)
}

# ---------------------------------------------------------------- VTT 解析 --

function ConvertFrom-VttFile {
    param([string]$Path)

    $lines = [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
    $cues  = New-Object System.Collections.Generic.List[object]

    # 時間軸行放寬比對：實際遇過的字幕檔會把 --> 逸出成 --&gt;，
    # 時間碼本身也可能殘缺（例如 00:35:100）。若不認得這種行，
    # 整行就會被當成字幕內容混進逐字稿。
    $stamp = '^\s*([\d:,\.]+)\s*-->\s*([\d:,\.]+)'
    $lastEnd = 0.0

    $i = 0
    while ($i -lt $lines.Count) {
        $line = $lines[$i] -replace '--&gt;', '-->'
        if ($line -match $stamp) {
            $start = ConvertTo-Seconds $Matches[1]
            $end   = ConvertTo-Seconds $Matches[2]
            # 壞掉的時間碼會讓 .srt 失序，這裡強制單調遞增
            if ($start -lt $lastEnd) { $start = $lastEnd }
            if ($end -le $start)     { $end   = $start + 2 }
            $lastEnd = $end
            $i++
            $body = New-Object System.Collections.Generic.List[string]
            while ($i -lt $lines.Count -and $lines[$i].Trim() -ne '') {
                $body.Add($lines[$i]); $i++
            }
            $cues.Add([pscustomobject]@{ Start = $start; End = $end; Lines = $body })
        }
        $i++
    }

    # YouTube 自動字幕會滾動重複前一行，這裡以最近 8 行為窗口去重
    $segs   = New-Object System.Collections.Generic.List[object]
    $recent = New-Object System.Collections.Generic.List[string]

    foreach ($cue in $cues) {
        $text = ''
        foreach ($raw in $cue.Lines) {
            $t = Clear-CueMarkup $raw
            if ($t -eq '') { continue }
            if ($recent.Contains($t)) { continue }
            $recent.Add($t)
            while ($recent.Count -gt 8) { $recent.RemoveAt(0) }
            $text = Join-Text $text $t
        }
        if ($text -ne '') {
            $segs.Add([pscustomobject]@{ Start = $cue.Start; End = $cue.End; Text = $text })
        }
    }
    return ,$segs.ToArray()
}

# ------------------------------------------------------------ 各式輸出格式 --

function New-RawParagraphs {
    param([object[]]$Segments, [double]$ChunkSeconds = 30)
    $paras = New-Object System.Collections.Generic.List[string]
    $buf = ''; $anchor = $null
    foreach ($s in $Segments) {
        if ($null -eq $anchor) { $anchor = $s.Start }
        $buf = Join-Text $buf $s.Text
        if (($s.End - $anchor) -ge $ChunkSeconds) {
            $paras.Add($buf); $buf = ''; $anchor = $null
        }
    }
    if ($buf -ne '') { $paras.Add($buf) }
    return ,$paras.ToArray()
}

function New-TidyParagraphs {
    <#
      補標點的兩個依據：字幕之間的停頓長度，以及累積字數。
      單靠停頓不夠用——人工字幕常常逐句相接、停頓恆為 0，
      這時就得改由字數決定逗號與句號的落點。
    #>
    param(
        [object[]]$Segments,
        [double]$CommaGap       = 0.25,   # 超過這個停頓就下逗號
        [double]$PeriodGap      = 0.80,   # 超過這個停頓就下句號
        [double]$ParagraphGap   = 1.80,   # 超過這個停頓就換段
        [int]   $CommaAfter     = 25,     # 安全網：太久沒有逗號就補一個
        [int]   $PeriodAfter    = 90,     # 安全網：太久沒有句號就斷句
        [int]   $ParagraphAfter = 240     # 段落超過這麼多字，在句尾換段
    )

    # 語言線索：句尾語氣詞收句，連接詞前後斷句，比單純數字數自然得多
    $stopTail  = '(了|嗎|呢|吧|啊|喔|囉|嘛|耶|哦|唷|謝謝)$'
    $commaTail = '(的話|然後|所以|但是|可是|因為|如果|其實|那麼|而且|甚至|例如|譬如|比如|以及|還有|同時|另外|不過|當然|就是說|第[一二三四五六七八九十]個)$'
    $leadHead  = '^(那|所以|但是|可是|然後|因為|如果|其實|不過|另外|而且|當然|我想|我覺得|接下來|最後|總之|也就是說|第[一二三四五六七八九十])'

    $paras = New-Object System.Collections.Generic.List[string]
    $buf = ''
    $sinceComma = 0
    $sinceStop  = 0
    $prevEnd = $null

    foreach ($s in $Segments) {
        $t = Clear-Filler $s.Text
        if ($t -eq '') { continue }

        $gap = 0.0
        if ($null -ne $prevEnd) { $gap = [double]$s.Start - [double]$prevEnd }
        $prevEnd = $s.End

        if ($buf -ne '') {
            $last   = $buf[$buf.Length - 1]
            $ended  = ('。！？.!?'.IndexOf($last)   -ge 0)
            $paused = ('，、；：,;:'.IndexOf($last) -ge 0)

            if (-not $ended -and -not $paused) {
                $stopHere =  ($gap -ge $PeriodGap) -or
                             (($buf -match $stopTail) -and $sinceStop -ge 10) -or
                             ($sinceStop -ge $PeriodAfter)
                $commaHere = ($gap -ge $CommaGap) -or
                             ($buf -match $commaTail) -or
                             (($t -match $leadHead) -and $sinceComma -ge 8) -or
                             ($sinceComma -ge $CommaAfter)

                if ($stopHere) {
                    $buf += '。'
                    $sinceStop = 0; $sinceComma = 0; $ended = $true
                }
                elseif ($commaHere) {
                    $buf += '，'
                    $sinceComma = 0
                }
            }
            elseif ($ended) { $sinceStop = 0; $sinceComma = 0 }

            # 換段：句子已收尾且段落夠長，或說話者明顯停頓
            if (($ended -and $buf.Length -ge $ParagraphAfter) -or $gap -ge $ParagraphGap) {
                $done = Close-Sentence $buf
                if ($done) { $paras.Add($done) }
                $buf = ''; $sinceComma = 0; $sinceStop = 0
            }
        }

        $buf = Join-Text $buf $t
        $sinceComma += $t.Length
        $sinceStop  += $t.Length
    }

    if ($buf -ne '') {
        $done = Close-Sentence $buf
        if ($done) { $paras.Add($done) }
    }
    return ,$paras.ToArray()
}

function New-TimecodeText {
    param([object[]]$Segments, [double]$ChunkSeconds = 20)
    $sb = New-Object System.Text.StringBuilder
    $buf = ''; $anchor = $null
    foreach ($s in $Segments) {
        if ($null -eq $anchor) { $anchor = $s.Start }
        $buf = Join-Text $buf $s.Text
        if (($s.End - $anchor) -ge $ChunkSeconds) {
            [void]$sb.AppendLine(('[{0}] {1}' -f (Format-Timestamp $anchor), $buf))
            [void]$sb.AppendLine()
            $buf = ''; $anchor = $null
        }
    }
    if ($buf -ne '' -and $null -ne $anchor) {
        [void]$sb.AppendLine(('[{0}] {1}' -f (Format-Timestamp $anchor), $buf))
    }
    return $sb.ToString()
}

function New-SrtText {
    param([object[]]$Segments)
    $sb = New-Object System.Text.StringBuilder
    $n = 1
    foreach ($s in $Segments) {
        [void]$sb.AppendLine($n)
        [void]$sb.AppendLine(('{0} --> {1}' -f (Format-Timestamp $s.Start 'srt'),
                                               (Format-Timestamp $s.End   'srt')))
        [void]$sb.AppendLine($s.Text)
        [void]$sb.AppendLine()
        $n++
    }
    return $sb.ToString()
}

function Write-Utf8File {
    param([string]$Path, [string]$Content)
    $enc = New-Object System.Text.UTF8Encoding($true)   # 含 BOM，記事本開啟不亂碼
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

function Get-SafeFileName {
    param([string]$Name, [int]$MaxLength = 80)
    $t = $Name -replace '[\\/:*?"<>|]', '_'
    $t = $t -replace '\s+', ' '
    $t = $t.Trim().TrimEnd('.')
    if ($t.Length -gt $MaxLength) { $t = $t.Substring(0, $MaxLength).Trim() }
    if ($t -eq '') { $t = 'transcript' }
    return $t
}

# ------------------------------------------------------------------ 字幕 --

function Get-SubtitleSegments {
    param([string]$VideoUrl, [string]$TempDir, [string[]]$CommonArgs)

    foreach ($mode in @('--write-subs', '--write-auto-subs')) {
        if ($mode -eq '--write-subs') { $label = '人工字幕' } else { $label = '自動字幕' }
        Write-Host ("  - 嘗試抓取{0}…" -f $label) -ForegroundColor DarkGray

        $ytArgs = @('--skip-download', $mode, '--sub-langs', $SubLangs,
                    '--sub-format', 'vtt/best', '--convert-subs', 'vtt',
                    '-o', (Join-Path $TempDir '%(id)s.%(ext)s')) + $CommonArgs + @($VideoUrl)
        Invoke-Native -Exe $script:YtDlp -Arguments $ytArgs -Quiet

        $vtts = @(Get-ChildItem -Path $TempDir -Filter '*.vtt' -ErrorAction SilentlyContinue)
        if ($vtts.Count -eq 0) { continue }

        # 依語言優先序挑一份
        $priority = @('zh-TW', 'zh-Hant', 'zh-Hant-TW', 'zh-tw', 'zh',
                      'zh-Hans', 'zh-CN', 'zh-cn', 'en', 'en-US')
        $best = $null; $bestRank = 9999
        foreach ($f in $vtts) {
            $lang = $f.Name -replace '^[^.]*\.', '' -replace '\.vtt$', ''
            $rank = $priority.IndexOf($lang)
            if ($rank -lt 0) { $rank = 500 }
            if ($rank -lt $bestRank) { $bestRank = $rank; $best = $f; $script:PickedLang = $lang }
        }

        $segs = ConvertFrom-VttFile $best.FullName
        if ($segs.Count -gt 0) {
            Write-Host ("  [完成] 取得{0}（{1}），共 {2} 個片段" -f $label, $script:PickedLang, $segs.Count) -ForegroundColor Green
            $script:SourceLabel = ('{0}（{1}）' -f $label, $script:PickedLang)
            return ,$segs
        }
    }
    return @()
}

function Get-WhisperSegments {
    param([string]$VideoUrl, [string]$TempDir, [string[]]$CommonArgs)

    if (-not $script:Python) { throw '找不到 Python，無法執行語音辨識。請先執行 setup.ps1。' }
    $pyScript = Join-Path $script:Root 'lib\transcribe.py'
    if (-not (Test-Path $pyScript)) { throw "找不到 $pyScript" }

    Write-Host '  - 下載音訊…' -ForegroundColor DarkGray
    $baseArgs = @('-f', 'bestaudio/best', '-x', '--audio-format', 'wav',
                  '--postprocessor-args', 'ffmpeg:-ar 16000 -ac 1',
                  '-o', (Join-Path $TempDir '%(id)s.%(ext)s')) + $CommonArgs + @($VideoUrl)
    if ($script:FFmpeg) {
        $baseArgs = @('--ffmpeg-location', (Split-Path -Parent $script:FFmpeg)) + $baseArgs
    }

    # YouTube 有時會對預設的播放器用戶端回 HTTP 403，換一個用戶端就能取得。
    # 這幾個是 yt-dlp 內建的選項，依序重試。非 YouTube 網站會忽略這個參數。
    $wav = $null
    $lastOutput = ''
    foreach ($client in @('', 'web_safari', 'mweb')) {
        $ytArgs = $baseArgs
        if ($client) {
            Write-Host ('    · 下載被拒，改用 {0} 播放器用戶端重試…' -f $client) -ForegroundColor DarkGray
            $ytArgs = @('--extractor-args', ('youtube:player_client={0}' -f $client)) + $baseArgs
        }

        $prev = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try { $lastOutput = (& $script:YtDlp @ytArgs 2>&1 | Out-String) }
        finally { $ErrorActionPreference = $prev }

        $wav = Get-ChildItem -Path $TempDir -Filter '*.wav' -ErrorAction SilentlyContinue |
               Select-Object -First 1
        if ($wav) { break }
    }

    if (-not $wav) {
        $detail = ($lastOutput -split "`n" | Where-Object { $_ -match 'ERROR|error' } |
                   Select-Object -First 1)
        if ($detail) { Write-Host ('    yt-dlp：{0}' -f $detail.Trim()) -ForegroundColor DarkGray }
        throw '音訊下載失敗，無法進行語音辨識。若影片需要登入，請加上 -CookiesFromBrowser edge。'
    }

    $json = Join-Path $TempDir 'whisper.json'
    Write-Host ('  - Whisper 辨識中（模型 {0}）…長片需要數分鐘至數十分鐘。' -f $Model) -ForegroundColor DarkGray

    $pyArgs = @($pyScript, '--audio', $wav.FullName, '--out', $json,
                '--model', $Model, '--compute-type', $ComputeType, '--traditional')
    if ($WhisperLang)   { $pyArgs += @('--lang', $WhisperLang) }
    if ($script:FFmpeg) { $pyArgs += @('--ffmpeg', $script:FFmpeg) }
    Invoke-Native -Exe $script:Python -Arguments $pyArgs
    if ($LASTEXITCODE -ne 0) { throw 'Whisper 辨識失敗。' }

    $data = Get-Content -Path $json -Raw -Encoding UTF8 | ConvertFrom-Json
    $segs = New-Object System.Collections.Generic.List[object]
    foreach ($s in $data.segments) {
        $segs.Add([pscustomobject]@{ Start = [double]$s.start; End = [double]$s.end; Text = $s.text })
    }
    $script:SourceLabel = ('Whisper 語音辨識（{0}）' -f $Model)
    Write-Host ('  [完成] 辨識完成，共 {0} 個片段' -f $segs.Count) -ForegroundColor Green
    return ,$segs.ToArray()
}

# ------------------------------------------------------------------ 主流程 --

function Invoke-OneVideo {
    param([string]$VideoUrl)

    $commonArgs = @('--no-warnings', '--no-playlist')
    if ($CookiesFromBrowser) { $commonArgs += @('--cookies-from-browser', $CookiesFromBrowser) }

    Write-Host ''
    Write-Host ('>> 處理：{0}' -f $VideoUrl) -ForegroundColor Cyan

    # 有些平台 yt-dlp 不認得，先把真正的影音位址解出來。
    # 後續交給 yt-dlp 的是 $mediaUrl，寫進逐字稿檔頭的仍是使用者給的原始網址。
    $resolved = Resolve-VideoSource -InputUrl $VideoUrl
    $mediaUrl = $resolved.Url
    if ($resolved.Note) { Write-Host ('  來源：{0}' -f $resolved.Note) -ForegroundColor DarkGray }

    # 暫存資料夾先開好：影片資訊改用 --print-to-file 寫檔再以 UTF-8 讀回，
    # 避免 PowerShell 以主控台編碼解讀 yt-dlp 的輸出，導致中文標題變亂碼
    $tmp = Join-Path $env:TEMP ('ytts_' + [guid]::NewGuid().ToString('N').Substring(0, 12))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null

    try {
        $sep = '@@@'
        $tpl = '%(id)s' + $sep + '%(title)s' + $sep + '%(duration)s' + $sep +
               '%(uploader)s' + $sep + '%(upload_date)s'
        $infoFile = Join-Path $tmp 'info.txt'
        $infoArgs = @('--skip-download', '--print-to-file', $tpl, $infoFile) +
                    $commonArgs + @($mediaUrl)
        Invoke-Native -Exe $script:YtDlp -Arguments $infoArgs -Quiet

        $info = $null
        if (Test-Path $infoFile) {
            $info = [System.IO.File]::ReadAllLines($infoFile, [System.Text.Encoding]::UTF8) |
                    Where-Object { $_.Trim() -ne '' } | Select-Object -First 1
            Remove-Item $infoFile -Force -ErrorAction SilentlyContinue
        }
        if (-not $info) { throw '無法讀取影片資訊，請確認網址是否正確、或影片是否需要登入。' }

        $f = $info -split [regex]::Escape($sep)
        $title    = $f[1]
        $duration = 0.0; [double]::TryParse($f[2], [ref]$duration) | Out-Null
        $uploader = $f[3]
        $upload   = $f[4]

        # 直接餵 m3u8 給 yt-dlp 時中繼資料是空的（標題會是 playlist），
        # 這時改用解析階段從原網頁取得的標題與站名。
        if ($resolved.Title) { $title = $resolved.Title }
        if ($resolved.Site -and (-not $uploader -or $uploader -eq 'NA')) { $uploader = $resolved.Site }
        foreach ($n in 'title', 'uploader', 'upload') {
            if ((Get-Variable $n -ValueOnly) -eq 'NA') { Set-Variable $n -Value '' }
        }
        if (-not $title) { $title = 'transcript' }

        $durText = if ($duration -gt 0) { Format-Timestamp $duration } else { '未提供（辨識後才知道）' }
        Write-Host ('  影片：{0}' -f $title)
        Write-Host ('  頻道：{0}    長度：{1}' -f $(if ($uploader) { $uploader } else { '未提供' }), $durText)

        $segs = @()
        if (-not $ForceWhisper) {
            $segs = Get-SubtitleSegments -VideoUrl $mediaUrl -TempDir $tmp -CommonArgs $commonArgs
        }
        if ($segs.Count -eq 0) {
            if (-not $ForceWhisper) {
                Write-Host '  - 這部影片沒有可用字幕，改用語音辨識。' -ForegroundColor Yellow
            }
            $segs = Get-WhisperSegments -VideoUrl $mediaUrl -TempDir $tmp -CommonArgs $commonArgs
        }
        if ($segs.Count -eq 0) { throw '未能取得任何逐字內容。' }

        # 組出各種版本
        $tidy   = New-TidyParagraphs -Segments $segs
        $rawPar = New-RawParagraphs  -Segments $segs
        $tcText = New-TimecodeText   -Segments $segs
        $srt    = New-SrtText        -Segments $segs

        if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
        $safe = Get-SafeFileName $title
        $base = Join-Path $OutDir $safe

        # 走語音辨識時，來源沒給長度也能從最後一個片段推得
        if ($duration -le 0 -and $segs.Count -gt 0) { $duration = [double](@($segs)[-1].End) }

        $meta = @(
            ('頻道：{0}' -f $(if ($uploader) { $uploader } else { '未提供' })),
            ('影片長度：{0}    上傳日期：{1}' -f (Format-Timestamp $duration),
                $(if ($upload) { $upload } else { '未提供' })),
            ('逐字稿來源：{0}    產生時間：{1}' -f $script:SourceLabel, (Get-Date -Format 'yyyy-MM-dd HH:mm'))
        )
        $header = (($meta + @($VideoUrl)) -join "`r`n") + "`r`n" + ('=' * 60) + "`r`n`r`n"

        Write-Utf8File ($base + '_逐字稿_整理版.txt') ($header + ($tidy   -join "`r`n`r`n"))
        Write-Utf8File ($base + '_逐字稿_原始.txt')   ($header + ($rawPar -join "`r`n`r`n"))
        Write-Utf8File ($base + '_逐字稿_時間碼.txt') ($header + $tcText)
        Write-Utf8File ($base + '.srt') $srt

        $chars = ($tidy -join '').Length
        Write-Host ('  [完成] 文字檔已輸出，整理版約 {0} 字' -f $chars) -ForegroundColor Green

        if (-not $NoDocx) {
            . (Join-Path $script:Root 'lib\New-TranscriptDocx.ps1')
            $pages = New-TranscriptDocx -OutPath ($base + '_逐字稿.docx') -Title $title `
                        -Url $VideoUrl -MetaLines $meta -TidyParagraphs $tidy -RawParagraphs $rawPar
            Write-Host ('  [完成] Word 檔已輸出，共 {0} 頁' -f $pages) -ForegroundColor Green
        }

        Write-Host ('  輸出資料夾：{0}' -f (Resolve-Path $OutDir)) -ForegroundColor Green
    }
    finally {
        if ($KeepTemp) {
            Write-Host ('  暫存檔保留於：{0}' -f $tmp) -ForegroundColor DarkGray
        } else {
            Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# ---------------------------------------------------------------------------

Initialize-Tools

$ok = 0; $fail = 0
foreach ($u in $Url) {
    try   { Invoke-OneVideo -VideoUrl $u; $ok++ }
    catch { Write-Host ('  [失敗] {0}' -f $_.Exception.Message) -ForegroundColor Red; $fail++ }
}

Write-Host ''
Write-Host ('全部完成：成功 {0} 部，失敗 {1} 部。' -f $ok, $fail) -ForegroundColor Cyan
