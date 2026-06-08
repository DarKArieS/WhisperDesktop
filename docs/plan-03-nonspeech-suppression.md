# 實作規劃 #03：非語音 token 抑制（Suppression）

> 對應 `feature-gap-vs-whisper-cpp.md` 第三項。
> 參考 whisper.cpp `whisper_process_logits` 與 OpenAI Whisper `tokenizer.non_speech_tokens`。

## 目標

原本 `sampleBest`（`ContextImpl.cpp`）取樣前只跳過 `sot/solm/not` 三個 token，
沒有完整抑制清單，容易吐出 `♪`、`「」`、`(( ))`、emoji/標點雜訊（尤其音樂段落會狂吐 `♪`）。
本次補上 whisper.cpp 的 `suppress_blank` 與 `suppress_nst`（non-speech tokens）。

## 設計

1. **抑制清單建一次**：`buildSuppressTokens()` 依詞表把非語音符號（含 `「」『』`、
   `♩♪♫♬♭♮♯`、各種括號/雙引號序列）的 token id 收集起來，每個符號同時查「本體」與
   「前綴空白」兩種形式，只有在詞表中為單一 token 時才加入；額外允許 `-`、`'`
   出現在字中但不在字首（抑制 ` -`、` '`）。去重後快取於 `suppress_tokens`。

2. **取樣前遮罩**：在 `sampleBest` 建好 `probs_id`（此時 `probs_id[i]` 仍 1:1 對應
   token id i，尚未排序）後、時間戳判斷與 top-K 之前，把要抑制的 token 機率設為
   `-INFINITY`（沿用本檔既有的 `-INFINITY` 遮罩慣例，機率已 softmax）。
   - `suppress_nst`：遮罩 `suppress_tokens`。
   - `suppress_blank` 且 `is_initial`：遮罩 `token_eot` 與空白 token（` `）。

3. **旗標控制**：新增 `eFullParamsFlags::SuppressBlank (0x400)`、
   `SuppressNonSpeech (0x800)`，bitfield 不改結構大小，無 ABI 風險；
   `fullDefaultParams` 預設兩者皆開（對齊 whisper.cpp 預設），GUI/PS 透過
   `fullDefaultParams` 自動沿用。

## 受影響檔案

| 檔案 | 變更 |
|---|---|
| `Whisper/API/sFullParams.h` | `eFullParamsFlags` 新增 `SuppressBlank` / `SuppressNonSpeech` |
| `Whisper/Whisper/ContextImpl.h` | 宣告 `buildSuppressTokens`、`suppress_tokens`、`token_space`、旗標布林 |
| `Whisper/Whisper/ContextImpl.cpp` | 實作 `buildSuppressTokens`、`sampleBest` 套用遮罩、`runFullImpl` 讀旗標並建表 |
| `Whisper/Whisper/ContextImpl.misc.cpp` | `fullDefaultParams` 預設開啟兩旗標 |
| `WhisperNet/API/Parameters.cs` | C# `eFullParamsFlags` 同步新增兩值 |

## 實作狀態（2026-06-07）

- 完成。native（`Whisper.dll`）與 managed（`WhisperNet.dll`）皆編譯通過。
- 僅 bitfield 旗標，無結構欄位變動、無跨 COM-light marshal ABI 風險。

## 驗收方式

- 轉錄含音樂/掌聲的音訊，確認不再大量輸出 `♪`、`(( ))`、`「」` 等非語音 token。
- 對照 whisper.cpp 同檔輸出，非語音雜訊應大致消失。
