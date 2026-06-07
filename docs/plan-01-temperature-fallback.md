# 實作規劃 #01：Temperature Fallback（溫度回退）

> 對應 `feature-gap-vs-whisper-cpp.md` 第一項。
> 參考 whisper.cpp `a8ec021`（2026-06-06）。

## 目標

以 whisper.cpp 的溫度回退機制，取代目前 `ContextImpl.cpp:775-866` 那段
「重複 >5 次就輸出 `repeat too many times! jump...` 並跳過」的粗糙 hack，
解決亂碼/重複/丟字問題。

---

## 關鍵前提（決定設計）

1. **GPU 已做完 softmax**：`WhisperContext::decode`（`WhisperContext.cpp:633`）
   回傳的 `probs` 是**機率**不是 logits；whisper.cpp 則對 logits 除溫度再 softmax
   （`whisper.cpp:6201-6205`）。

2. **不必改 shader**：用 softmax 平移不變性，純 CPU 還原溫度分佈：
   `tempered_prob[i] ∝ probs[i]^(1/T)`（等價 `softmax(log(probs)/T)`，加性常數抵消）。
   T=0 時走原本 greedy。

3. **判斷指標用 entropy，非 compression ratio**：whisper.cpp 靠 **avg_logprob** 與
   **最後 32 token 的 entropy**（`whisper.cpp:6627-6648`、`7548`），不需 gzip。

4. **no_speech token 找得到**：本專案 `token_solm`（多語版 50362，原始碼註 `// ??`）
   正好落在 `<|nospeech|>`（whisper.cpp `token_nosp`）位置。
   `no_speech_prob = probs[token_solm]`，每個 seek 的第一次 decode（未加溫度）算一次。

5. **KV cache 可安全重試**：每個 seek 先 `encode` 一次（cross-attn `kvCross` 固定）；
   重試只需從 `n_past=0` 重跑 decode，覆寫同一塊 self-attn KV 區
   （`decodeLayer` offset 含 `n_past`）。**重試間不可呼叫 `clearState`**，encode 不重算。

---

## 流程設計（每個 seek）

```
encode(mel, seek)                          // 一次
建 prompt（prompt_init + 前文 context）     // 一次，先不寫回 prompt_past
no_speech_prob = 第一次 decode 後 probs[token_solm]   // 每 seek 一次

for t in {0.0, 0.2, 0.4, 0.6, 0.8, 1.0}:   // 由 temperature / temperature_inc 產生
    tokens, result_len, seek_delta, failed = decodeSegment(t)
    avg_logprob = mean(plog over result tokens)
    entropy     = entropyOfLast32(tokens)
    if result_len > 32 && entropy < entropy_thold(2.4): failed = true
    if 最後一個溫度: 接受; break
    if failed || (avg_logprob < logprob_thold(-1.0)
                  && no_speech_prob < no_speech_thold(0.6)):
        continue                            // 升溫重試
    接受; break

is_no_speech = no_speech_prob > 0.6 && avg_logprob < -1.0
commit：寫回 prompt_past、result_all、觸發 new_segment_callback
        （is_no_speech 則只前進 seek、不輸出文字）
seek += seek_delta
```

### 取樣
- T=0：現有 `sampleBest`（argmax）。
- T>0：新增 `sampleWithTemperature`，沿用 `sampleBest` 的時間戳遮罩規則，
  以 `std::discrete_distribution`（`<random>` 已 include）從 `probs[i]^(1/T)` 抽樣；
  `plog = log(該分佈中選中 token 的機率)`。

---

## 分階段

- **Phase 1（核心）**：新增參數 + 把 `runFullImpl` 內層迴圈抽成 `decodeSegment`，
  加溫度回退 + entropy 重試守衛 + `sampleWithTemperature`，**移除** repeat hack
  （`ContextImpl.cpp:775-866`）。
- **Phase 2**：`no_speech_prob` 略過輸出、`best_of`（每溫度抽多條取最佳分數）。

---

## 受影響檔案

| 檔案 | 變更 |
|---|---|
| `Whisper/API/sFullParams.h` | 結構**尾端**新增 `temperature / temperature_inc / entropy_thold / logprob_thold / no_speech_thold` |
| `Whisper/Whisper/ContextImpl.misc.cpp` | `fullDefaultParams` 設預設（0.0 / 0.2 / 2.4 / -1.0 / 0.6） |
| `Whisper/Whisper/ContextImpl.h` | 宣告 `decodeSegment`、`sampleWithTemperature`、`std::mt19937` 成員 |
| `Whisper/Whisper/ContextImpl.cpp` | 重構 `runFullImpl`、新增函式、移除 repeat hack |

## 相容性風險

- `sFullParams` 跨 COM-light marshal 到 C#（WhisperNet）。欄位**只能加在尾端**，
  並須同步 C# struct，否則 ABI 錯位。Phase 1 可先讓原生端用預設值生效（不動 GUI）。
- 重試只在 t=0 未被接受時發生，正常音訊零額外開銷。

## 實作狀態（2026-06-07）

- **Phase 1：完成、已驗證、native + managed 皆編譯通過。**
  - `sampleWithTemperature`、`decodeSegment`、entropy 守衛、`runFullImpl` 重構、移除 repeat hack。
  - 參數加在 `sFullParams`（native）與 `sFullParams.cs`（C#）尾端；`fullDefaultParams` 回填預設，GUI 自動啟用。
- **Phase 2：完成、編譯通過。**
  - `best_of`：t>0 時抽 `best_of` 條獨立樣本，取分數（avg_logprob，非 failed 優先）最高者；預設 2，append 在 native/C# 結構尾端。
  - `is_no_speech`：`no_speech_prob > 0.6 && avg_logprob < -1.0` 的段落不輸出、也不寫入 prompt_past。
  - 尾端短音訊（`seek+500 >= seek_end`）清空 prompt_past，減少結尾幻聽（對應 whisper.cpp:7046-7051）。
  - 註：本專案 native `greedy` 結構僅有 `n_past`，故 `best_of` 以頂層欄位 append，而非放進 greedy 子結構。

- **Phase 3：時間戳規則（修「重複輸出 + 時間戳不準」的真正根因）。編譯通過。**
  - 根因：本專案 sampler 缺了 whisper.cpp `whisper_process_logits` 的時間戳規則。最關鍵是
    **時間戳必須非遞減**（whisper.cpp:6330-6338）：未限制時模型可能吐出較早的時間戳，
    使 `seek` 幾乎不前進 → 下個窗格大量重疊 → 同一段文字被反覆轉錄（重複），時間戳也漂移。
    原碼只有解碼迴圈裡一個弱的「break 整段」檢查，擋不住。
  - 新增 `applyDecodingRules()`，在每次取樣前對機率切片做遮罩（機率已 softmax，遮罩 = 設 0）：
    1. 抑制控制/語言 token（sot、not、prev、solm、translate、transcribe、語言）。
    2. **時間戳成對**（whisper.cpp:6298-6317）：強制 `[ts] text [ts]` 結構。
    3. **時間戳非遞減**（whisper.cpp:6330-6338）：遮罩低於目前 `seek_delta/2` 的時間戳。
  - `no_speech_prob` 在套用規則前先讀（因規則會遮罩 token_solm）。純 native，無 ABI 變動。

## 驗收方式

- 找一段會觸發目前 repeat hack 的音訊，確認改後不再出現
  `repeat too many times! jump...`，且文字連續無大量重複。
- 比對同檔在 whisper.cpp（同模型）輸出，段落切點與內容大致一致。
