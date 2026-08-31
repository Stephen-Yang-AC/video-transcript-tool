# -*- coding: utf-8 -*-
"""以 faster-whisper 將音訊檔轉為帶時間碼的逐字稿 JSON。

由 Get-YouTubeTranscript.ps1 呼叫，只在影片沒有字幕可抓時才會用到。
輸出格式：{"language": "zh", "duration": 123.4,
           "segments": [{"start": 0.0, "end": 3.2, "text": "..."}, ...]}

註：本機的 Windows 應用程式控制原則（Smart App Control）會封鎖 PyAV 的原生
DLL，導致 `import av` 失敗。faster-whisper 只在「輸入是檔案路徑」時才會用到
PyAV 解碼音訊；若直接傳入 numpy 陣列則完全用不到。因此這裡改由 ffmpeg 解碼，
並在必要時放一個 av 替身模組讓 faster_whisper 得以載入。這不會降低辨識品質，
也不需要更動任何系統安全設定。
"""
import argparse
import json
import os
import subprocess
import sys
import types

# Windows 上沒有開發人員模式時無法建立符號連結，HuggingFace 會每次跳警告。
# 快取照常運作，只是會多佔一點磁碟空間，這裡直接關掉這則訊息。
os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")

import numpy as np

SAMPLE_RATE = 16000


def ensure_av_importable():
    """PyAV 若被系統原則封鎖，就放一個空殼模組頂替。"""
    try:
        import av  # noqa: F401
        return True
    except Exception as exc:
        sys.modules["av"] = types.ModuleType("av")
        print("[提示] PyAV 無法載入（%s），改用 ffmpeg 解碼音訊。" % exc.__class__.__name__,
              file=sys.stderr, flush=True)
        return False


def decode_audio(path, ffmpeg="ffmpeg"):
    """用 ffmpeg 把任意音訊解碼成 16 kHz 單聲道 float32 陣列。"""
    cmd = [ffmpeg, "-nostdin", "-loglevel", "error", "-i", path,
           "-f", "f32le", "-ac", "1", "-ar", str(SAMPLE_RATE), "-"]
    try:
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
    except FileNotFoundError:
        raise RuntimeError("找不到 ffmpeg，請先執行 setup.ps1。")
    except subprocess.CalledProcessError as exc:
        raise RuntimeError("ffmpeg 解碼失敗：%s" % exc.stderr.decode("utf-8", "replace").strip())

    audio = np.frombuffer(proc.stdout, dtype=np.float32)
    if audio.size == 0:
        raise RuntimeError("音訊解碼結果為空，請確認檔案是否正常。")
    return audio


def get_converter(enabled):
    if not enabled:
        return None
    try:
        from opencc import OpenCC
        return OpenCC("s2twp")
    except Exception:
        print("[提示] 未安裝 opencc，略過簡轉繁後處理。", file=sys.stderr)
        return None


def build_parser():
    p = argparse.ArgumentParser(description="faster-whisper 語音辨識")
    p.add_argument("--audio", required=True, help="輸入音訊檔")
    p.add_argument("--out", required=True, help="輸出 JSON 路徑")
    p.add_argument("--model", default="large-v3-turbo", help="模型名稱")
    p.add_argument("--lang", default=None, help="語言代碼，如 zh、en；省略則自動偵測")
    p.add_argument("--device", default="cpu", help="cpu 或 cuda")
    p.add_argument("--compute-type", default="int8", help="int8 / int8_float32 / float32")
    p.add_argument("--beam-size", type=int, default=5)
    p.add_argument("--ffmpeg", default="ffmpeg", help="ffmpeg 執行檔路徑")
    p.add_argument("--prompt", default="以下是繁體中文的演講逐字稿。",
                   help="initial_prompt，用來引導模型輸出繁體中文")
    p.add_argument("--traditional", action="store_true",
                   help="若已安裝 opencc，將結果轉為臺灣繁體用字")
    return p


def main():
    args = build_parser().parse_args()

    ensure_av_importable()
    try:
        from faster_whisper import WhisperModel
    except ImportError as exc:
        print("[錯誤] 無法載入 faster-whisper：%s" % exc, file=sys.stderr)
        print("       請先執行 setup.ps1。", file=sys.stderr)
        return 2

    print("[Whisper] 解碼音訊…", file=sys.stderr, flush=True)
    audio = decode_audio(args.audio, args.ffmpeg)
    duration = len(audio) / float(SAMPLE_RATE)
    print("[Whisper] 音訊長度 %.1f 秒" % duration, file=sys.stderr, flush=True)

    print("[Whisper] 載入模型 %s（%s / %s）…首次執行需下載模型，請耐心等候。"
          % (args.model, args.device, args.compute_type), file=sys.stderr, flush=True)
    model = WhisperModel(args.model, device=args.device, compute_type=args.compute_type)

    print("[Whisper] 開始辨識…", file=sys.stderr, flush=True)
    segments, info = model.transcribe(
        audio,
        language=args.lang,
        beam_size=args.beam_size,
        vad_filter=True,
        vad_parameters={"min_silence_duration_ms": 500},
        condition_on_previous_text=False,   # 避免長影片落入重複迴圈
        initial_prompt=args.prompt or None,
    )

    cc = get_converter(args.traditional)
    out = []
    for seg in segments:
        text = (seg.text or "").strip()
        if not text:
            continue
        if cc is not None:
            text = cc.convert(text)
        out.append({"start": round(seg.start, 3), "end": round(seg.end, 3), "text": text})
        pct = (seg.end / duration * 100) if duration else 0
        print("[%5.1f%%] %s" % (pct, text[:60]), file=sys.stderr, flush=True)

    payload = {"language": getattr(info, "language", args.lang),
               "duration": duration,
               "segments": out}
    with open(args.out, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)

    print("[Whisper] 完成，共 %d 個片段。" % len(out), file=sys.stderr, flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
