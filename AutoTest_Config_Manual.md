# AutoTest_Config.xlsx 填寫說明書

> **適用對象**：需要新增、修改或維護自動化測試項目的工程師  
> **適用程式**：`general_auto_engine.py`（SemiAutoRunner 自動測試引擎）  
> **最後更新**：2026-06-06

---

## 目錄

1. [檔案總覽](#1-檔案總覽)
2. [Sheet ①：Item_Config（測項總覽與控制）](#2-sheet-item_config測項總覽與控制)
3. [Sheet ②：Formula_Macro（變數計算與公式定義）](#3-sheet-formula_macro變數計算與公式定義)
4. [Sheet ③：Layout_Mapping（步驟排程與報表對位）](#4-sheet-layout_mapping步驟排程與報表對位)
5. [三張表的跨表關聯約束](#5-三張表的跨表關聯約束)
6. [Meas_Types 關鍵字完整對照表](#6-meas_types-關鍵字完整對照表)
7. [單位自動轉換規則](#7-單位自動轉換規則)
8. [實戰範例：從零新增一個測項](#8-實戰範例從零新增一個測項)
9. [常見問題與除錯 (FAQ)](#9-常見問題與除錯-faq)
10. [附錄：現有測項清單速查](#10-附錄現有測項清單速查)
11. [AI 代理人介面 (AI Agent Bridge) 操作指引](#11-ai-代理人介面-ai-agent-bridge-操作指引)

---

## 1. 檔案總覽

### 1.1 定位與用途

`AutoTest_Config.xlsx` 是自動化測試引擎的**唯一外部控制檔**，負責將 Python 程式碼與測試邏輯完全分離。透過編輯此檔案，使用者不需修改任何 Python 腳本即可：

| 能力 | 說明 |
| :--- | :--- |
| 新增 / 停用測項 | 在 `Item_Config` 表增刪行或切換 `Enable` 欄位 |
| 更換示波器設定 | 修改 `.set` 檔案路徑 |
| 調整負載計算邏輯 | 在 `Formula_Macro` 表修改 Python 公式 |
| 調整報表對位 | 在 `Layout_Mapping` 表修改圖片錨點與資料儲存格 |

### 1.2 工作表結構（MECE 分層）

本檔案精確包含三張工作表，各司其職、互不重疊：

```
AutoTest_Config.xlsx
├─ Item_Config       ← 「做什麼」：測項清單、開關、報表 Sheet、示波器設定檔
├─ Formula_Macro     ← 「用什麼」：變數定義、從報表讀取基準值、Python 計算公式
└─ Layout_Mapping    ← 「怎麼做」：執行步驟、提示詞、截圖位置、量測擷取方式
```

### 1.3 資料流概念圖

```
Item_Config          Formula_Macro           Layout_Mapping
┌──────────┐     ┌────────────────┐     ┌──────────────────┐
│Test_Item  │────▶│Test_Item       │     │Test_Item         │
│Enable     │     │Var_Name ──────────────▶Load_Variable    │
│Setup_File │     │Spec_Cell       │     │Phase_Step        │
│Report_Sheet────────────────────────────▶Picture_Anchor   │
│Vout_Cell  │     │Python_Eval_    │     │Vpp_Data_Cell     │
│           │     │  Formula       │     │Meas_Types        │
└──────────┘     └────────────────┘     └──────────────────┘
```

---

## 2. Sheet ①：Item_Config（測項總覽與控制）

### 2.1 角色

定義**要測試哪些項目**、是否啟用、對應的報告工作表名稱、以及示波器設定檔路徑。引擎啟動時只會掃描此表來決定執行範圍。

### 2.2 欄位定義

| # | 欄位名稱 | 必填 | 資料型態 | 說明 | 合法值 / 格式 | 範例 |
| :---: | :--- | :---: | :--- | :--- | :--- | :--- |
| A | **Test_Item** | ✅ | 字串 | 測項唯一識別 ID。必須與 `Formula_Macro` 和 `Layout_Mapping` 中的 `Test_Item` **完全一致**（含大小寫、底線、句點）。 | 自訂，建議格式 `{編號}.{名稱}_{類型}` | `2.Ripple_and_Noise_POL` |
| B | **Enable** | ✅ | 字串 | 測項開關。引擎只執行標記為 `YES` 的項目。 | `YES` 或 `NO`（必須大寫） | `YES` |
| C | **Report_Sheet** | ✅ | 字串 | 測試報告 Excel 中，波形圖與數據要回填的**目標工作表名稱**。必須是報告檔 `.xlsm` 中已存在的 Sheet 名稱。 | 報告檔中的 Sheet 名稱 | `2.Ripple_and_Noise_POL` |
| D | **Setup_File** | ✅ | 字串 | 示波器 `.set` 設定檔的路徑。支援兩種格式（見下方說明）。 | 絕對路徑或示波器內部路徑 | `C:\...\2.Ripple&Noise.set` |
| E | **Vout_Cell** | ❌ | 字串 | 輸出電壓所在的報告儲存格，供註記或未來擴充使用。目前程式未直接讀取。 | Excel 儲存格地址，或留空 | `D27` |

### 2.3 Setup_File 路徑格式

| 格式 | 說明 | 範例 |
| :--- | :--- | :--- |
| **本機絕對路徑** | 檔案存在於測試電腦上。引擎會逐行讀取檔案內容並透過 SCPI 推送給示波器。 | `C:\TestEnv\ScopeSets\2.Ripple&Noise.set` |
| **示波器內部路徑** | 檔案存在於示波器上（如 USB 隨身碟）。引擎會直接下達 `RECAll:SETUp` 命令。 | `E:/DPO4054_POL_set/Ripple.set` |

### 2.4 填寫規則與限制

- `Test_Item` 不可重複（同一個 ID 只能出現一次）
- `Enable` 只接受大寫 `YES` / `NO`，不接受 `yes`、`True`、`1` 等其他寫法
- `Report_Sheet` 必須是報告檔中已存在的工作表名稱，否則執行時會拋出 `KeyError`
- `Setup_File` 若為本機路徑，檔案必須實際存在，否則引擎會嘗試以示波器內部路徑處理

---

## 3. Sheet ②：Formula_Macro（變數計算與公式定義）

### 3.1 角色

定義每個測項需要的**計算變數**。引擎會從報告檔的 `VR SPEC` 工作表中讀取基準數值，再套用此表定義的 Python 公式即時運算。運算結果會被 `Layout_Mapping` 的 `Load_Variable` 引用，用於提示操作人員設定電子負載。

### 3.2 欄位定義

| # | 欄位名稱 | 必填 | 資料型態 | 說明 | 合法值 / 格式 | 範例 |
| :---: | :--- | :---: | :--- | :--- | :--- | :--- |
| A | **Test_Item** | ✅ | 字串 | 對應 `Item_Config` 中的測項名稱，精確匹配。同一個 `Test_Item` 可以有多筆（多個變數）。 | 與 `Item_Config.Test_Item` 一致 | `2.Ripple_and_Noise_POL` |
| B | **Var_Name** | ✅ | 字串 | 變數名稱。此名稱會被 `Layout_Mapping.Load_Variable` 引用。同一測項內不可重複。 | 自訂，建議用蛇形命名法 | `Light_load` |
| C | **Spec_Cell** | 條件必填 | 字串 | 報告檔 `VR SPEC` 工作表中，要讀取的基準數值所在儲存格。若公式不需要透過 `Spec[]` 字典取值，可填 `N/A`。 | Excel 儲存格地址（如 `I36`）或 `N/A` | `I36` |
| D | **Python_Eval_Formula** | ✅ | 字串 | Python 運算式，由引擎透過 `eval()` 動態執行。 | 合法的 Python 表達式 | `float(Spec['I36']) * 0.1` |

### 3.3 Python_Eval_Formula 可用變數

公式在 `eval()` 執行時，可使用以下兩個預注入的變數：

| 變數名 | 型態 | 說明 | 用法範例 |
| :--- | :--- | :--- | :--- |
| `Spec` | `dict` | 字典。鍵為本表 `Spec_Cell` 中定義的儲存格地址，值為從報告 `VR SPEC` 工作表該儲存格讀取的原始數值。若該儲存格為空，預設填入 `0.0`。 | `float(Spec['I36'])` |
| `ws` | `xlwings.Sheet` | 報告檔 `VR SPEC` 工作表的 xlwings Sheet 物件，可直接呼叫 `.range()` 等 API 存取任意儲存格。 | `str(ws.range('G40').value)` |

### 3.4 常見公式寫法

| 場景 | 公式範例 | 說明 |
| :--- | :--- | :--- |
| 直接取值 | `float(Spec['I36'])` | 讀取 I36 儲存格數值並轉為浮點數 |
| 比例計算 | `float(Spec['I36']) * 0.1` | 輕載 = 滿載電流 × 10% |
| 字串拼接提示 | `str(ws.range('G40').value) + ' to ' + str(ws.range('I40').value)` | 組合成 `"12.5 to 15.0"` 格式的文字提示 |
| 多儲存格運算 | `(float(Spec['I36']) + float(Spec['H36'])) / 2` | 取兩個值的平均 |

### 3.5 填寫規則與限制

- `Spec_Cell` 欄位若填寫有效儲存格地址，引擎會在執行公式前先將該儲存格的值讀入 `Spec` 字典
- `Spec_Cell` 填 `N/A` 或留空時，公式仍可正常執行，但不能在公式中引用不存在的 `Spec` 鍵，否則會拋出 `KeyError`
- 公式若要透過 `ws` 物件存取儲存格，不需在 `Spec_Cell` 欄位列出（因為 `ws` 可以直接存取任何儲存格）
- 公式計算結果若為數值，引擎會自動 `round()` 到小數點後兩位
- 公式計算結果若為字串，則直接保留原始格式

---

## 4. Sheet ③：Layout_Mapping（步驟排程與報表對位）

### 4.1 角色

定義每個測項的**逐步執行腳本**。引擎會依照 `Phase_Step` 順序，逐步提示操作人員、等待按鍵、擷取示波器波形、回填數據與圖片至報告。

### 4.2 欄位定義

| # | 欄位名稱 | 必填 | 資料型態 | 說明 | 合法值 / 格式 | 範例 |
| :---: | :--- | :---: | :--- | :--- | :--- | :--- |
| A | **Test_Item** | ✅ | 字串 | 對應 `Item_Config` 中的測項名稱。同一測項會有多行（每個 Phase 一行）。 | 與 `Item_Config.Test_Item` 一致 | `4.FSW_POL` |
| B | **Phase_Step** | ✅ | 整數 | 階段編號，從 `1` 起連續遞增。引擎會按此編號排序執行。 | 正整數 (1, 2, 3...) | `1` |
| C | **Prompt_Name** | ✅ | 字串 | 該階段在命令列上顯示的提示名稱，告知操作人員當前測試條件。 | 自訂文字 | `輕載 (0.1*Imax)` |
| D | **Load_Variable** | ✅ | 字串 | 參照 `Formula_Macro.Var_Name` 中算出的變數名稱。引擎會在螢幕上顯示該變數的計算結果（安培），提示測試人員調整電子負載機。 | 與 `Formula_Macro.Var_Name` 一致 | `Light_load` |
| E | **Picture_Anchor** | ✅ | 字串 | 波形截圖在報告工作表中的貼上範圍。引擎會自動將圖片縮放並貼齊此範圍。 | Excel 範圍格式 `{左上}:{右下}` | `B28:K49` |
| F | **Vpp_Data_Cell** | ✅ | 字串 | 量測數據要回填的 Excel 儲存格。多個儲存格用**半形逗號**分隔，數量必須與 `Meas_Types` 中的項目一一對應。 | 儲存格地址，逗號分隔 | `K61, K62` |
| G | **Meas_Types** | ✅ | 字串 | 定義從示波器擷取的量測數據類型。支援多種關鍵字語法（詳見[第 6 節](#6-meas_types-關鍵字完整對照表)）。多個項目用**半形逗號**分隔。 | 關鍵字或 `MEASn:關鍵字` | `MEAS3:MAX` |

### 4.3 欄位互動關係

`Vpp_Data_Cell` 與 `Meas_Types` 之間存在**嚴格的位置對應**關係：

```
Vpp_Data_Cell:  K61,       K62
                 │           │
                 ▼           ▼
Meas_Types:   MAXimum,   MINImum
```

- 逗號分隔的項目數量必須**完全相同**
- 第 1 個 `Meas_Types` 項目的結果 → 寫入第 1 個 `Vpp_Data_Cell`
- 第 2 個 `Meas_Types` 項目的結果 → 寫入第 2 個 `Vpp_Data_Cell`
- 以此類推

### 4.4 填寫規則與限制

- `Phase_Step` 必須從 `1` 開始，同一測項內遞增且不可跳號
- `Load_Variable` 必須存在於 `Formula_Macro` 表中，否則引擎會顯示 `0.0` 作為負載提示
- `Picture_Anchor` 格式必須為 `左上角:右下角`（如 `B28:K49`），不可只填單一儲存格
- `Vpp_Data_Cell` 和 `Meas_Types` 的逗號分隔項目數量必須一致

---

## 5. 三張表的跨表關聯約束

以下整理所有跨表欄位的參照關係及約束條件：

| 來源欄位 | 目標欄位 | 關聯性質 | 違反時的後果 |
| :--- | :--- | :--- | :--- |
| `Item_Config.Test_Item` | `Formula_Macro.Test_Item` | 一對多（一個測項可有多個變數） | 找不到變數，所有 `Load_Variable` 顯示 `0.0` |
| `Item_Config.Test_Item` | `Layout_Mapping.Test_Item` | 一對多（一個測項可有多個 Phase） | 找不到 Layout，該測項將跳過所有階段 |
| `Formula_Macro.Var_Name` | `Layout_Mapping.Load_Variable` | 多對一（多個 Phase 可引用同一變數） | `Load_Variable` 找不到對應 `Var_Name`，引擎顯示 `0.0` |
| `Item_Config.Report_Sheet` | 報告檔 `.xlsm` 中的 Sheet | 必須精確匹配 | 拋出 `KeyError` 並中斷執行 |

### 5.1 跨表一致性檢查清單

在修改設定檔前後，建議逐項確認：

- [ ] `Item_Config` 中每個 `Test_Item` 在 `Formula_Macro` 中至少有一筆記錄
- [ ] `Item_Config` 中每個 `Test_Item` 在 `Layout_Mapping` 中至少有一筆記錄
- [ ] `Layout_Mapping` 中每個 `Load_Variable` 在同測項的 `Formula_Macro.Var_Name` 中存在
- [ ] `Item_Config` 中的 `Report_Sheet` 是報告檔中已存在的 Sheet 名稱
- [ ] `Layout_Mapping` 中每列的 `Vpp_Data_Cell` 與 `Meas_Types` 項目數量一致

---

## 6. Meas_Types 關鍵字完整對照表

`Meas_Types` 欄位決定引擎從示波器擷取何種量測數據。支援以下關鍵字：

### 6.1 語法格式

| 語法 | 說明 | 範例 |
| :--- | :--- | :--- |
| `MEASn:關鍵字` | **精準指定**：明確指定讀取示波器 MEAS n 的特定統計值 | `MEAS3:MAX` |
| `關鍵字` | **自動編號**：依逗號位置自動對應 MEAS 1, 2, 3... | `MAXimum, MINImum` |

### 6.2 關鍵字對照表

| 關鍵字 | 擷取方式 (SCPI 命令) | 適用場景 | 單位轉換 |
| :--- | :--- | :--- | :--- |
| `MAX` / `MAXimum` | `MEASUrement:MEASn:MAXimum?` | Vpp 最大值（Ripple、Switching 等） | V → mV（×1000） |
| `MIN` / `MINImum` | `MEASUrement:MEASn:MINImum?` | Vpp 最小值（Dynamic Load 等） | V → mV（×1000） |
| `VAL` | `MEASUrement:MEASn:VALue?` | 即時量測值 | V → mV（×1000） |
| `VAL_FREQ` / 含 `FREQ` 的關鍵字 | `MEASUrement:MEASn:VALue?` → ÷ 1,000,000 | 切換頻率 (FSW) | Hz → MHz → 寫入 `"{值} kHz"` |
| `TIME_CURSOR` / 含 `CURSOR` 的關鍵字 | `CURSor:VBArs:DELTa?` | Jitter Cursor 差值 | s → ns（×10⁹），寫入 `"{值} ns"` |
| `VAL_WIDTH` / 含 `WIDTH` 的關鍵字 | `MEASUrement:MEASn:VALue?` | Jitter 脈寬量測 | s → ns（×10⁹），寫入 `"{值} ns"` |

> **注意**：關鍵字的匹配是**不區分大小寫**的，引擎會先將關鍵字轉為大寫再比對。

### 6.3 關鍵字判斷優先順序

引擎在解析 `Meas_Types` 時，按以下優先順序判斷：

```
1. 關鍵字包含 "CURSOR" → 使用 CURSor:VBArs:DELTa? 查詢
2. 關鍵字包含 "MIN"    → 使用 MEASn:MINImum? 查詢
3. 關鍵字包含 "VAL"    → 使用 MEASn:VALue? 查詢
   3a. 同時包含 "FREQ" → 額外除以 1,000,000
4. 其餘（含 "MAX"）    → 使用 MEASn:MAXimum? 查詢（預設行為）
```

---

## 7. 單位自動轉換規則

引擎在將量測值寫入報告時，會根據 `Meas_Types` 自動套用不同的單位轉換：

| Meas_Types / 測項條件 | 擷取原始單位 | 轉換公式 | 寫入報告格式 | 範例 |
| :--- | :--- | :--- | :--- | :--- |
| `FREQ` | Hz (示波器回傳) | ÷ 1,000,000 → MHz → × 1000 → kHz | `"{數值:.1f} kHz"` (字串) | `"1200.0 kHz"` |
| `Jitter` 測項（或包含 `TIME`/`WIDTH`/`CURSOR`） | 秒 (s) | × 10⁹ → ns | 純數值（ns） | `985.6`（報表格式會顯示 `985.6ns`） |
| `Dynamic_Load` 測項（或包含 `VOLT`/`RAW`） | 伏特 (V) | 直接寫入（不乘 1000） | 純數值（V） | `1.805`（即 1.805 V） |
| 其他（Ripple 等 mV 類） | 伏特 (V) | × 1000 → mV | 純數值（mV） | `15.3`（即 15.3 mV） |

> **重要**：
> - `6.Jitter` 測項（包含 Jitter Cursor 與 Pulse Width / MEAS1:VAL）皆為時間量測，引擎會自動以 **`× 10⁹` 轉換為奈秒（ns）純數值** 寫入報告儲存格，由 Excel 報表原本設定的儲存格自訂格式（如 `####.0"ns"`）呈現，不會再有少乘 1,000,000 的問題。
> - `3.Dynamic_Load` 因為量測的是輸出電壓 Vout Max / Min（單位是 V），引擎會自動以**原始純數值 (V)** 寫入，不會乘以 1000。
> - Ripple（紋波）測試則會依標準將 V 轉為 mV（×1000）寫入。

---

## 8. 實戰範例：從零新增一個測項

以下以新增 `7.OCP_POL`（過電流保護測試）為例，示範如何在三張表中各自新增記錄。

### Step 1：Item_Config 新增一行

| Test_Item | Enable | Report_Sheet | Setup_File | Vout_Cell |
| :--- | :--- | :--- | :--- | :--- |
| `7.OCP_POL` | `YES` | `7.Overcurrent_protection_POL` | `C:\...\DPO4054_POL_set\7.OCP.set` | `D27` |

### Step 2：Formula_Macro 定義變數

| Test_Item | Var_Name | Spec_Cell | Python_Eval_Formula |
| :--- | :--- | :--- | :--- |
| `7.OCP_POL` | `ocp_threshold` | `I36` | `float(Spec['I36']) * 1.2` |

### Step 3：Layout_Mapping 定義步驟

| Test_Item | Phase_Step | Prompt_Name | Load_Variable | Picture_Anchor | Vpp_Data_Cell | Meas_Types |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `7.OCP_POL` | `1` | `OCP 觸發 (1.2×Imax)` | `ocp_threshold` | `B28:K49` | `K50` | `MEAS1:MAX` |

### Step 4：建立批次檔（選用）

```batch
@echo off
title OCP Semi-Auto Test
cd /d "%~dp0"

echo ==============================================
echo       Running OCP Semi-Auto Test
echo ==============================================
echo.

.venv\Scripts\python.exe general_auto_engine.py "7.OCP_POL"

echo.
pause
```

---

## 9. 常見問題與除錯 (FAQ)

### Q1：執行腳本時直接略過了我的測項？

**原因與排查**：
- ✅ 檢查 `Item_Config` 表的 `Enable` 欄位是否填寫了**大寫** `YES`
- ✅ 如果是透過批次檔啟動（傳入 `specific_test` 參數），確認參數字串與 `Test_Item` 完全一致
- ✅ 確認 Excel 儲存格沒有多餘的空白或不可見字元

### Q2：為什麼測出來的頻率數值大了一千倍？

**原因**：程式預設會將量測值 × 1000（V → mV）。  
**解法**：確保 `Meas_Types` 中包含 `FREQ` 關鍵字（例如 `MEAS1:VAL_FREQ`），引擎才會啟用頻率專用的轉換邏輯（÷ 1,000,000 再 × 1000 → kHz 字串）。

### Q3：回填到 Excel 的資料錯位了？

**原因**：`Vpp_Data_Cell` 與 `Meas_Types` 的逗號分隔項目不一致。  
**解法**：確保兩者的項目數量和順序一一對應。例如：
```
Vpp_Data_Cell:  K61, K62
Meas_Types:     MAXimum, MINImum
→ MAXimum 的結果填入 K61，MINImum 的結果填入 K62
```

### Q4：公式執行時出現 `KeyError`？

**原因**：`Python_Eval_Formula` 中引用了 `Spec['XX']`，但 `Spec_Cell` 欄位沒有定義 `XX` 這個儲存格。  
**解法**：在 `Formula_Macro` 中為同一測項新增一行，`Spec_Cell` 填入需要的儲存格地址（如 `I36`），或改用 `ws.range('XX').value` 語法直接存取。

### Q5：示波器設定檔載入失敗？

**排查步驟**：
1. 確認 `Setup_File` 路徑格式正確（Windows 路徑用反斜線 `\`）
2. 若為本機路徑，確認檔案確實存在於該位置
3. 若為示波器內部路徑（如 `E:/...`），確認 USB 隨身碟已插入示波器
4. 檢查 `.set` 檔案是否為純文字格式且編碼正確（UTF-8 或 ASCII）

### Q6：報告儲存失敗（檔案被鎖定）？

**原因**：報告檔 `.xlsm` 正被其他 Excel 程式或程序佔用。  
**解法**：關閉所有手動開啟的報告檔副本，只保留引擎自動開啟的那一份。

### Q7：`ws.range()` 與 `Spec[]` 有什麼差別？

| 特性 | `Spec['XX']` | `ws.range('XX').value` |
| :--- | :--- | :--- |
| 速度 | 較快（已預載入記憶體） | 較慢（即時存取 Excel COM） |
| 前置條件 | 必須在 `Spec_Cell` 中定義 | 不需要，可存取任意儲存格 |
| 空值處理 | 自動填入 `0.0` | 回傳 `None` |
| 適用場景 | 數值計算（需要 float 轉換） | 文字提示、跨欄位拼接 |

---

## 10. 附錄：現有測項清單速查

### 10.1 Item_Config 現有測項

| Test_Item | Enable | Report_Sheet | Setup_File 檔名 |
| :--- | :---: | :--- | :--- |
| `2.Ripple_and_Noise_POL` | YES | `2.Ripple_and_Noise_POL` | `2.Ripple&Noise.set` |
| `3.Dynamic_Load_POL` | YES | `3.Dynamic_Load_POL` | `3.Dynamic.set` |
| `4.FSW_POL` | YES | `4.FSW under normal mode_POL` | `4.FSW.set` |
| `5.Switching_HS_Rise` | YES | `5.Switching waveforms_POL` | `5.VDS_HS_Rise.set` |
| `5.Switching_HS_Fall` | YES | `5.Switching waveforms_POL` | `5.VDS_HS_Fall.set` |
| `5.Switching_LS_Fall` | YES | `5.Switching waveforms_POL` | `5.VDS_LS_Fall.set` |
| `5.Switching_LS_Rise` | YES | `5.Switching waveforms_POL` | `5.VDS_LS_Rise.set` |
| `6.Jitter_POL` | YES | `6.Jitter_POL` | `6.Jitter_COT_mode.set` |

### 10.2 各測項的 Phase 結構摘要

| Test_Item | Phase 數 | Meas_Types 類型 | 特殊說明 |
| :--- | :---: | :--- | :--- |
| `2.Ripple_and_Noise_POL` | 3 | `MEAS3:MAX` | 輕 / 中 / 滿載 Ripple |
| `3.Dynamic_Load_POL` | 2 | `MAXimum, MINImum` | 兩組 Transient 條件，每組回填兩個值 |
| `4.FSW_POL` | 3 | `MEAS1:VAL_FREQ` | 頻率量測，結果以 kHz 字串寫入 |
| `5.Switching_HS_Rise` | 3 | `MAXimum, MINImum` | 高側 MOSFET 上升沿 |
| `5.Switching_HS_Fall` | 3 | `MAXimum, MINImum` | 高側 MOSFET 下降沿 |
| `5.Switching_LS_Fall` | 3 | `MAXimum, MINImum` | 低側 MOSFET 下降沿 |
| `5.Switching_LS_Rise` | 3 | `MAXimum, MINImum` | 低側 MOSFET 上升沿 |
| `6.Jitter_POL` | 3 | `TIME_CURSOR, VAL_WIDTH` | Jitter 量測，結果以 ns 字串寫入 |

### 10.3 負載條件常用變數

| Var_Name | 計算方式 | 意義 |
| :--- | :--- | :--- |
| `Full_load` | `float(Spec['I36'])` | 滿載電流 (Imax) |
| `Mid_load` | `float(Spec['H36'])` | 中載電流 (TDC) |
| `load_cond_1` | `str(ws.range('G40').value) + ' to ' + str(ws.range('I40').value)` | Dynamic Load 條件 1 (文字提示) |
| `load_cond_2` | `str(ws.range('G41').value) + ' to ' + str(ws.range('I41').value)` | Dynamic Load 條件 2 (文字提示) |

---

## 11. AI 代理人介面 (AI Agent Bridge) 操作指引

系統提供標準的 AI 代理人橋接模組（`ai_agent_bridge.py`），為未來多模態 AI（如 Gemini Vision）與自主測試代理人預留專屬 API 接口。

### 11.1 核心功能
1. **視覺快照介面 (`get_screen_image` / `get_screen_base64`)**：
   - 提取示波器即時畫面（PIL Image 或 Base64 編碼字串），供多模態視覺模型進行波形分析、觸發檢測與數值判讀。
2. **SCPI 儀器控制介面 (`send_scpi_command` / `query_scpi`)**：
   - 供 AI Agent 依據判讀結果動態調整時基、垂直刻度、觸發電位與量測槽。
3. **流程調度與確認 (`confirm_step` / `abort_test`)**：
   - 在全自動測試模式下，AI 代理人可在判定波形穩定後自動觸發步驟確認，實現無人值守測試閉環。

### 11.2 呼叫範例
```python
from ai_agent_bridge import AIAgentBridge
from general_auto_engine import AutoTestEngine

engine = AutoTestEngine()
ai_bridge = AIAgentBridge(engine=engine)

# 1. 取得當前示波器畫面 Base64
img_b64 = ai_bridge.get_screen_base64()

# 2. 供 AI 決策後發送 SCPI 指令
ai_bridge.send_scpi_command("TRIGGER:A:LEVEL 1.2")

# 3. 確認步驟並繼續
ai_bridge.confirm_step()
```

