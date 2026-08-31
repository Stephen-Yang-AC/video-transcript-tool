# 以 Word COM 產出逐字稿 .docx（標楷體 + Times New Roman、頁碼、整理版與原始版各一節）
# 由 Get-YouTubeTranscript.ps1 以點號載入（dot-source）後呼叫。

function New-TranscriptDocx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $OutPath,
        [Parameter(Mandatory = $true)][string] $Title,
        [string]   $Url             = '',
        [string[]] $MetaLines       = @(),
        [string[]] $TidyParagraphs  = @(),
        [string[]] $RawParagraphs   = @()
    )

    $word   = $null
    $doc    = $null
    $sel    = $null
    $footer = $null
    $fRange = $null

    # 記下呼叫前既有的 Word 進程，收尾時才分得出哪個是本函式開的。
    # 絕不能動到使用者自己開著的 Word。
    $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible       = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Add()

        # 版面：上下 2.54 cm、左右 3.17 cm
        $doc.PageSetup.TopMargin    = 72
        $doc.PageSetup.BottomMargin = 72
        $doc.PageSetup.LeftMargin   = 90
        $doc.PageSetup.RightMargin  = 90

        $sel = $word.Selection

        function Write-Para {
            param([string]$Text, [int]$Size = 14, [bool]$Bold = $false,
                  [int]$Align = 0, [int]$SpaceAfter = 6, [single]$FirstLineIndent = 0)
            $sel.ParagraphFormat.Alignment       = $Align   # 0 = 左、1 = 置中
            $sel.ParagraphFormat.SpaceAfter      = $SpaceAfter
            $sel.ParagraphFormat.LineSpacingRule = 5        # wdLineSpaceMultiple
            $sel.ParagraphFormat.LineSpacing     = 18       # 1.5 倍行高
            $sel.ParagraphFormat.FirstLineIndent = $FirstLineIndent
            $sel.Font.Size = $Size
            $sel.Font.Bold = [int]$Bold
            if ($Text) { $sel.TypeText($Text) }
            $sel.TypeParagraph()
        }

        # 標題
        Write-Para -Text $Title -Size 18 -Bold $true -Align 1 -SpaceAfter 12

        # 影片資訊
        foreach ($m in $MetaLines) {
            if ($m) { Write-Para -Text $m -Size 11 -Align 1 -SpaceAfter 2 }
        }
        if ($Url) { Write-Para -Text $Url -Size 11 -Align 1 -SpaceAfter 6 }
        Write-Para -Text '' -Size 11 -SpaceAfter 12

        # 一、整理版
        if ($TidyParagraphs.Count -gt 0) {
            Write-Para -Text '一、整理版逐字稿' -Size 16 -Bold $true -SpaceAfter 10
            foreach ($p in $TidyParagraphs) {
                if ($p) { Write-Para -Text $p -Size 14 -SpaceAfter 10 -FirstLineIndent 28 }
            }
        }

        # 二、原始逐字稿
        if ($RawParagraphs.Count -gt 0) {
            $sel.InsertBreak(7)   # wdPageBreak
            Write-Para -Text '二、原始逐字稿（未經潤飾，供精確引述之用）' -Size 16 -Bold $true -SpaceAfter 10
            foreach ($p in $RawParagraphs) {
                if ($p) { Write-Para -Text $p -Size 13 -SpaceAfter 8 -FirstLineIndent 26 }
            }
        }

        # 頁碼（頁尾置中）
        $footer = $doc.Sections.Item(1).Footers.Item(1)   # wdHeaderFooterPrimary
        $fRange = $footer.Range
        $fRange.ParagraphFormat.Alignment = 1
        $fRange.Font.Size = 10
        $doc.Fields.Add($fRange, -1, 'PAGE', $true) | Out-Null

        # 全文套用字型：中文標楷體、英數 Times New Roman
        $doc.Content.Font.NameFarEast = '標楷體'
        $doc.Content.Font.NameAscii   = 'Times New Roman'
        $doc.Content.Font.NameOther   = 'Times New Roman'
        $footer.Range.Font.NameFarEast = '標楷體'
        $footer.Range.Font.NameAscii   = 'Times New Roman'

        $dir = Split-Path -Parent $OutPath
        if ($dir -and -not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $doc.SaveAs2($OutPath, 16)          # wdFormatDocumentDefault = .docx
        $pages = $doc.ComputeStatistics(2)  # wdStatisticPages
        $doc.Close(0)
        $doc = $null
        return $pages
    }
    finally {
        if ($doc) { try { $doc.Close(0) } catch { } }

        # 先釋放所有子物件的 COM 參考再結束 Word。只要還有 RCW 沒放掉，
        # Quit() 之後 Word 仍可能滯留成沒有視窗的 WINWORD.EXE，
        # 一路累積並占住輸出資料夾（會導致該資料夾無法改名或刪除）。
        foreach ($o in @($fRange, $footer, $sel, $doc)) {
            if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
        }
        $fRange = $null; $footer = $null; $sel = $null; $doc = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()

        if ($word) {
            try { $word.Quit(0) } catch { }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
            $word = $null
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()

        # 保險：若仍有本函式開出來、且沒有視窗的 Word 殘留，才結束它。
        # 兩個條件缺一不可，確保不會誤關使用者自己開著的文件。
        Start-Sleep -Milliseconds 300
        foreach ($p in @(Get-Process WINWORD -ErrorAction SilentlyContinue)) {
            if ($before -notcontains $p.Id -and -not $p.MainWindowTitle) {
                try { Stop-Process -Id $p.Id -Force -ErrorAction Stop } catch { }
            }
        }
    }
}
