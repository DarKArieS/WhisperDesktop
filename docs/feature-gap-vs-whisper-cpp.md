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

### 1. Temperature fallback（溫度回退）— ✅ 已實作
whisper.cpp 在一段解碼失敗時（觸發 `entropy_thold` /
`logprob_thold` / `no_speech_thold`）會自動以遞增溫度重試
（0.0→0.2→…→1.0），是它穩定不亂碼的核心機制。

已補上完整的 temperature fallback：`sFullParams` 新增 `temperature`、`temperature_inc`、
`entropy_thold`、`logprob_thold`、`no_speech_thold`、`best_of` 欄位（預設值與 whisper.cpp
相同：0.0 / 0.2 / 2.4 / -1.0 / 0.6 / 5），並與 `WhisperNet/Internal/sFullParams.cs` 的
ABI 對齊。`runFullImpl` 改成：每個 seek 先 encode 一次，再以遞增溫度重複呼叫新的
`decodeSegment`，依品質門檻（最後 32 token 的 entropy、平均 logprob、no-speech 機率）決定
是否回退。`sampleBest` 支援 `temperature > 0` 時依 `p^(1/T)` 分布抽樣（GPU 已 softmax，
故用 log-prob 還原溫度）。原本 `"repeat too many times! jump..."` 與
`"do not find segment!"` 兩個會丟字的 hack 已移除。

### 2. Beam search — 宣告了但沒實作
`sFullParams.h` 的 `eSamplingStrategy::BeamSearch` 註解寫著
`// TODO: not implemented yet!`。`runFullImpl` 只有 greedy
（`ContextImpl.cpp:681`）。也沒有 `best_of` / 多 decoder。

### 3. 非語音 token 抑制（suppression）— ✅ 已實作
whisper.cpp 有 `suppress_blank`、`suppress_nst`（non-speech tokens）、
`suppress_regex`。本專案原本的 `sampleBest`（`ContextImpl.cpp`）只跳過
`sot/solm/not` 三個 token，沒有完整抑制清單，容易吐出 `♪`、emoji、標點雜訊。

已補上 `suppress_blank` 與 `suppress_nst`（含 `♩♪♫♬♭♮♯`、`「」『』`、括號序列等），
旗標 `SuppressBlank` / `SuppressNonSpeech` 預設開啟。詳見 `plan-03-nonspeech-suppression.md`。
（`suppress_regex` 暫未做。）

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
| temperature / entropy 門檻 | ✅ | ✅ **已實作**（temperature fallback，見上方第 1 項） |

---

## 建議優先順序（針對維護）

1. ~~**Temperature fallback + entropy/logprob 門檻**~~ — ✅ 已完成，
   已解掉「重複就跳過丟字」的根本問題。
2. **量化模型載入**（`WhisperModel.cpp` 的 ftype 分支 + 對應 dequant compute shader）
   — 解 VRAM 與相容性。
3. **重新啟用並改用 DTW 時間戳** — 目前是註解掉的死碼。
4. **Beam search** — 補完那個 TODO。
5. **神經 VAD（Silero）** 整合到擷取與離線流程。
