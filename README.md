# video-transcript-tool

**影片逐字稿工具** · <https://github.com/Stephen-Yang-AC/video-transcript-tool>

把線上影片轉成逐字稿。優先抓取影片既有的字幕，沒有字幕才動用本機語音辨識。

## 支援的網站

| 網站 | 取得方式 |
|------|----------|
| YouTube | 人工字幕 → 自動字幕 → 語音辨識 |
| 教育大市集 `market.cloud.edu.tw` | 解析頁面取得 HLS 串流；該平台不提供字幕，一律走語音辨識。若該教材其實掛在 YouTube，會自動改走 YouTube 並優先抓字幕 |
| 其他 | yt-dlp 支援上千個網站，多數影音網址可直接貼上試試 |

---

## 一、安裝

只需執行一次。已安裝的項目會自動略過，重複執行是安全的。

```powershell
.\setup.ps1
```

會安裝：yt-dlp（抓字幕與音訊）、FFmpeg（音訊轉檔）、Python 3.12、
faster-whisper（本機語音辨識）、opencc（簡轉繁）。

若想先把 Whisper 模型下載好，避免第一次辨識時等待：

```powershell
.\setup.ps1 -PrefetchModel large-v3-turbo
```

---

## 二、基本用法

一律用 `逐字稿.cmd` 啟動器即可，它不會被 PowerShell 的執行原則擋下。

**不帶參數：雙擊執行**

在檔案總管裡雙擊 `逐字稿.cmd`，貼上網址按 Enter。

**一支影片**

```powershell
.\逐字稿.cmd "https://www.youtube.com/watch?v=XXXXXXXXXXX"
```

**多支影片**：網址以空白分隔，會依序處理。

```powershell
.\逐字稿.cmd "https://youtu.be/AAAA" "https://youtu.be/BBBB"
```

**加上參數**：網址寫在前，參數接在後。

```powershell
.\逐字稿.cmd "https://youtu.be/XXXX" -ForceWhisper -NoDocx
```

網址請務必用雙引號包住——YouTube 網址常含 `&`，在命令列裡不加引號會被截斷。

<details>
<summary>也可以直接呼叫主程式</summary>

```powershell
.\Get-VideoTranscript.ps1 -Url "https://www.youtube.com/watch?v=XXXXXXXXXXX"
```

用法與啟動器相同，只是若出現「因為這個系統上已停用指令碼執行」，
得先自行處理執行原則，見第七節。

</details>

---

## 三、輸出檔案

全部存到 `輸出\` 資料夾，檔名取自影片標題：

| 檔案 | 用途 |
|------|------|
| `_逐字稿_整理版.txt` | 補標點、分段、去口頭禪，可直接閱讀引用 |
| `_逐字稿_原始.txt` | 忠實呈現字幕原文，不做任何改寫，供精確引述 |
| `_逐字稿_時間碼.txt` | 每段前加 `[hh:mm:ss]`，方便回頭核對原片段落 |
| `.srt` | 標準字幕檔，可掛回影片或匯入剪輯軟體 |
| `_逐字稿.docx` | 標楷體＋Times New Roman、含頁碼；第一節整理版、第二節原始版 |

---

## 四、常用參數

| 參數 | 說明 |
|------|------|
| `-OutDir <路徑>` | 改變輸出資料夾，預設為工具目錄下的 `輸出\` |
| `-NoDocx` | 只產文字檔，不開 Word |
| `-ForceWhisper` | 略過字幕、直接用語音辨識（字幕品質太差時使用） |
| `-Model <名稱>` | Whisper 模型，預設 `large-v3-turbo` |
| `-WhisperLang <代碼>` | 語音語言，預設 `zh`；留空字串則自動偵測 |
| `-CookiesFromBrowser <瀏覽器>` | 會員限定或需登入的影片，例如 `edge`、`chrome` |
| `-KeepTemp` | 保留暫存檔（字幕原檔、音訊），供除錯用 |

範例——影片字幕是簡體或機器字幕，想改用語音辨識重跑：

```powershell
.\逐字稿.cmd "https://youtu.be/XXXX" -ForceWhisper
```

---

## 五、運作方式

1. **先抓人工字幕。** 品質最好，多數正式訪談與講座都有。
2. **再抓自動字幕。** YouTube 機器產生，會滾動重複前一行，工具會自動去重。
3. **都沒有才用 Whisper。** 下載音訊後在本機辨識，全程不上傳雲端。

一小時的影片，走字幕路徑約十幾秒完成；走語音辨識則需要數十分鐘（本機為 CPU 運算）。

### Whisper 模型怎麼選

| 模型 | 速度 | 中文品質 | 建議場合 |
|------|------|----------|----------|
| `medium` | 較快 | 中上 | 只需抓大意 |
| `large-v3-turbo` | 中等 | 好 | **預設值**，多數情況最划算 |
| `large-v3` | 慢 | 最好 | 需要逐字精確、且願意等 |

---

## 六、已知限制

- **整理版的標點是規則式推估**，依停頓長度、句尾語氣詞與連接詞落點判斷，
  約八成位置合理，但不等於人工潤稿。需要精確引用時請以「原始版」為準；
  需要通順成文時，把整理版交給 Claude 再潤一次會更好。
- **來源字幕本身的錯字不會被修正。** 工具只做整理，不改內容。
- **部分影片的字幕檔有瑕疵**（實際遇過上傳者誤把 SRT 片段貼進字幕文字、
  時間軸被 HTML 逸出成 `--&gt;`）。工具會濾掉這類標記並修正錯序的時間碼，
  但該處的字幕斷句可能較粗。
- **無 NVIDIA 顯示卡**，語音辨識走 CPU，速度受限。

---

## 七、疑難排解

**「因為這個系統上已停用指令碼執行」**
Windows 用戶端的預設執行原則是 `Restricted`，會擋掉所有 .ps1。三種處理方式：

1. **用 `逐字稿.cmd` 啟動器**（推薦）。它以 Bypass 模式呼叫主程式，
   只影響那一次呼叫，不更動任何系統設定。
2. **只放行目前這個視窗**：在同一個視窗先執行
   `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`，
   關掉視窗就失效。
3. **永久放行自己的帳號**（不需系統管理員權限，會變更安全設定，請自行斟酌）：
   `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`。
   `RemoteSigned` 的意思是本機自己寫的指令碼可以跑，從網路下載的則必須有數位簽章。

**「找不到 yt-dlp」**
執行 `.\setup.ps1`。若剛裝完仍找不到，關掉 PowerShell 視窗重開一個。

**Whisper 報 `DLL load failed ... 應用程式控制原則已封鎖此檔案`**
本機的 Windows Smart App Control 會封鎖 PyAV 的原生 DLL。
工具已內建繞法：改由 FFmpeg 解碼音訊，不需要更動任何系統安全設定。
若仍出現此訊息，確認 `setup.ps1` 有把 FFmpeg 裝好。

**「音訊下載失敗，無法進行語音辨識」**
YouTube 有時會對下載請求回 HTTP 403。工具會自動換三種播放器用戶端重試
（預設 → `web_safari` → `mweb`），多數情況會自行解決，畫面上會看到重試訊息。
若三種都失敗，多半是該影片需要登入，加上 `-CookiesFromBrowser edge` 再試。

**影片需要登入才能看**
加上 `-CookiesFromBrowser edge`（或 `chrome`），借用瀏覽器既有的登入狀態。

**Word 檔沒產生**
確認電腦已安裝 Microsoft Word。若沒有，加 `-NoDocx` 只產文字檔。

---

## 八、檔案結構

```
video-transcript-tool\
├─ 逐字稿.cmd                   啟動器（雙擊即用，繞過執行原則）
├─ setup.ps1                    安裝相依套件
├─ Get-VideoTranscript.ps1    主程式
├─ README.md                    本說明
├─ lib\
│  ├─ transcribe.py             faster-whisper 語音辨識
│  └─ New-TranscriptDocx.ps1    Word COM 產出 docx
└─ 輸出\                        逐字稿存放處
```
