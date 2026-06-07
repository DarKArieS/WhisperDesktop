# WhisperDesktop 與最新 whisper.cpp 的功能差距分析

> 比較日期：2026-06-07
> 對照 whisper.cpp 版本：`a8ec021`（2026-06-06）(C:\Users\aries\VSProject\whisper.cpp\src\whisper.cpp)
> 本專案分支：`aries/adjust`（已含 large-v3、repeat filter、diarize）

## 背景定位

WhisperDesktop（Const-me/Whisper）**不是** whisper.cpp 的封裝，而是 2022–23 年用
DirectCompute（HLSL compute shader）從頭重寫的 GPU 推論引擎。它的解碼邏輯凍結在
**早期 whisper.cpp（`whisper_sample_best` 純 greedy 時代）**。

由於後端完全不同，CUDA / Metal / Flash-attention / OpenVINO / CoreML 這類
「whisper.cpp 的後端功能」對本專案不適用，不算缺失。真正的差距在
**解碼品質、模型格式、時間戳、VAD**。

---

## 一、解碼品質（最大缺口）

### 1. Temperature fallback（溫度回退）— 最關鍵缺失
whisper.cpp 在一段解碼失敗時（觸發 `compression_ratio_threshold` /
`logprob_threshold` / `no_speech_threshold`）會自動以遞增溫度重試
（0.0→0.2→…→1.0），是它穩定不亂碼的核心機制。

本專案 **完全沒有 temperature 概念**（`Whisper/API/sFullParams.h` 連欄位都沒有）。
取而代之的是 `Whisper/Whisper/ContextImpl.cpp:807-866` 一個粗糙 hack：偵測到同一段
文字重複 >5 次就直接塞一句 `"repeat too many times! jump..."` 跳過，會直接丟字。

### 2. Beam search — 宣告了但沒實作
`sFullParams.h` 的 `eSamplingStrategy::BeamSearch` 註解寫著
`// TODO: not implemented yet!`。`runFullImpl` 只有 greedy
（`ContextImpl.cpp:681`）。也沒有 `best_of` / 多 decoder。

### 3. 非語音 token 抑制（suppression）
whisper.cpp 有 `suppress_blank`、`suppress_nst`（non-speech tokens）、
`suppress_regex`。本專案的 `sampleBest`（`ContextImpl.cpp:126`）只跳過
`sot/solm/not` 三個 token，沒有完整抑制清單，容易吐出 `♪`、emoji、標點雜訊。

### 4. no_speech 偵測
沒有 `no_speech_thold` 與 `no_speech_prob`，無法判斷整段是靜音/音樂而跳過。

---

## 二、模型格式

### 5. 量化模型完全不支援
`Whisper/Whisper/WhisperModel.cpp:301-310`（及 hybrid 版 390-399）只認
`ftype==0`→FP32，其餘一律當 FP16。**Q4_0 / Q5_1 / Q8_0 等量化 GGML 都無法載入**。
whisper.cpp 早已支援，量化能省一半以上 VRAM。

### 6. 仍是舊 `ggml` 格式
`WhisperModel.cpp:443` 檢查 magic `0x67676d6c`（"ggml"）。能讀標準 fp16 模型，
但不支援後來的量化擴充頭。

### 7. large-v3-turbo / distil-whisper
turbo 結構上（layer 數從 header 讀）**可能**能載入（已補 n_mels=128），但未驗證；
distil 系列未支援。

---

## 三、時間戳

### 8. DTW token 級時間戳未啟用
whisper.cpp 用 alignment heads + DTW 做精準字級時間戳。本專案只有舊的能量 /
voice_length 啟發式 `expComputeTokenLevelTimestamps`（`ContextImpl.cpp:274`），
而且**在主迴圈裡是被註解掉的**（`ContextImpl.cpp:844-849`、`874-879`）。
`max_len` 自動斷句（`wrapSegment`）同樣被註解掉，等於 `TokenTimestamps` /
`max_len` 旗標目前無效。

---

## 四、VAD

### 9. 無神經網路 VAD（Silero）
whisper.cpp 近期整合 Silero VAD 模型，先切出語音段再轉錄，大幅提速並減少幻聽。
本專案只有 `Whisper/Whisper/voiceActivityDetection.cpp` 那個 2009 年能量門檻演算法，
且只用於即時麥克風擷取（README 自承即時延遲達 5–10 秒）。

---

## 五、其他功能差距

| 功能 | whisper.cpp | 本專案 |
|---|---|---|
| Grammar / GBNF 約束解碼 | ✅ | ❌ |
| tinydiarize (tdrz) 語者 token | ✅ | 只有自製雙聲道 diarize（`ContextImpl.diarize.cpp`） |
| 字級信心 / karaoke 上色輸出 | ✅ | ❌ |
| 自動語言偵測 | ✅ | ✅ **已實作**（`detectLanguage` `ContextImpl.cpp:62`；README 第 139 行「not implemented」已過時） |
| temperature / entropy 門檻 | ✅ | ❌ |

---

## 建議優先順序（針對維護）

1. **Temperature fallback + compression-ratio/logprob 門檻** — CP 值最高，
   直接解掉現在「重複就跳過丟字」的根本問題。
2. **量化模型載入**（`WhisperModel.cpp` 的 ftype 分支 + 對應 dequant compute shader）
   — 解 VRAM 與相容性。
3. **重新啟用並改用 DTW 時間戳** — 目前是註解掉的死碼。
4. **Beam search** — 補完那個 TODO。
5. **神經 VAD（Silero）** 整合到擷取與離線流程。
