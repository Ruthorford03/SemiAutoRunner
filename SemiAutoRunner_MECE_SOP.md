# SemiAutoRunner 半自動電源驗證 MECE SOP

| 文件欄位 | 內容 |
| --- | --- |
| 適用系統 | `SemiAutoRunner`、Tektronix DPO4054、Power Validation Report |
| 適用對象 | 測試執行者、設定維護者、程式維護者 |
| 文件版本 | v1.3 |
| 盤點日期 | 2026-09-15 |
| 依據 | 現有 `.bat`、Python、`AutoTest_Config.xlsx`、示波器 `.set`、報告 `.xlsm` 的靜態核對；另以舊版活頁簿範本作為操作邏輯與需求參考 |
| 驗證範圍 | 已驗證檔案結構、設定關聯、報告儲存格及程式控制流；2026-09-14 下午的現有 `.bat` 實機操作完成狀態由操作者提供；舊版活頁簿僅做唯讀分析，未執行巨集或連線儀器 |

## 1. 目的與適用邊界

本系統把「示波器設定、人工調整負載、波形擷取、量測值讀回、Excel 報告回填」串成半自動流程。操作人員仍負責接線、供電、設定電子負載、確認波形穩定及判斷量測是否合理；程式不會自動控制電子負載，也不會替代硬體安全判斷。

本 SOP 依 MECE 分為六個互不重疊的管理面：

1. 執行環境：Windows、Python、Excel、VISA、網路。
2. 輸入資料：報告檔與 `VR SPEC`。
3. 測試設定：`AutoTest_Config.xlsx`。
4. 儀器設定：各測項 `.set`。
5. 執行控制：`.bat` 與 `general_auto_engine.py`。
6. 輸出驗收：量測值、截圖、報告儲存與異常復原。

## 2. 現行狀態與後續維護確認事項

2026-09-14 下午，操作者已完成當日要求的全部 `.bat` 測試，並確認現有 `.venv` 可正常使用。本節保留靜態盤點所見的維護確認事項，但不得在未比對當日報告與示波器畫面的情況下直接修改現行環境。

| 等級 | 現況 | 直接影響 | 放行條件 |
| --- | --- | --- | --- |
| 已確認 | 2026-09-14 下午已完成當日要求的全部 `.bat` 測試 | 現行批次入口與 `.venv` 可供操作者使用 | 保存當日報告、測試條件與版本作為基準 |
| 維護確認 | 靜態 Config 中 OCP 的 `Formula_Macro.Var_Name` 兩列都叫 `Quick Current`，Layout 則使用 `Quick Current Rise`、`Slow Current Rise` | 程式碼路徑可能取得預設 `0.0`，但不否定操作者已完成的實測 | 先比對當日 Console、負載設定與報告，再決定是否需要修改 |
| 維護確認 | 靜態 Config 中 Jitter 的第二個 `Meas_Types` 是未指定槽號的 `VAL_WIDTH`，程式會依位置讀 `MEAS2`；`.set` 的 Negative Width 在 `MEAS1` | 需確認當日 Pulse Width 是否來自預期量測槽 | 先核對當日示波器畫面與報告值，不直接改 Config |
| 安全提醒 | 示波器連線失敗時，程式會進入模擬模式；按 Enter 仍會把 `0.05` 經單位換算後寫入報告 | 離線假資料可能被儲存成正式結果 | 現行 SOP 規定連線失敗立即中止，不按 Enter |
| 維護確認 | 先前檔案快照曾出現 `9.91E+37` 類示波器無效值 | 無效量測可能被格式化成極大 kHz 字串 | 以 2026-09-14 當日 FSW 報告為準確認是否仍發生 |
| 維護確認 | OCP Config 使用 `MEAS2:MAX`，而 `.set` 中 CH2 Maximum 類型配置在 MEAS3 | 量測槽與報告語意需由實測證據確認 | 比對當日示波器 Measurement 列、Console 與 K70/K98 |
| 範圍說明 | 8.SCP 與 10.Sequence 沒有專用 `.bat`，不屬於本次「現有 `.bat` 全部完成」的範圍 | 不影響本次已完成的批次測試紀錄 | 未來若要納入半自動流程，再補齊 Formula、Layout 與批次入口 |
| 資訊 | 現有 `.venv` 已由操作者確認可透過雙擊 `.bat` 正常執行 | 現況是可用的 Golden Environment | 凍結環境；不得因文件盤點而重建、升級或安裝套件 |
| 中 | `pyproject.toml` 的 dependencies 為空 | 無法單靠專案檔描述目前可用環境 | 只記錄現況；不得在正式環境自動同步或安裝依賴 |
| 中 | 路徑及示波器 IP 寫死 | 搬資料夾、換報告或換儀器後會失效或寫錯檔 | 完成第 12 節的路徑變更檢查 |
| 中 | Config 中部分中文 Prompt 已成為 `����` | 操作提示不可讀 | 以正確 Unicode 文字重新輸入並確認 Excel 儲存後仍正常 |

維護確認事項的處理原則是「先保存證據、再判斷是否變更」。操作者已完成的正式結果應依當日 Test Plan、示波器畫面及報告驗收，不以靜態程式檢查單獨推翻。

## 3. 系統組成與唯一責任

| 元件 | 唯一責任 | 日常是否修改 |
| --- | --- | --- |
| `啟動_*.bat` | 選定一個 Test Item，切到正確目錄後呼叫 Python | 新增或改名測項時 |
| `general_auto_engine.py` | 串接 Config、Excel、示波器、人工確認、量測與回填 | 只有功能或錯誤處理變更時 |
| `AutoTest_Config.xlsx` | 定義測項、變數公式、Phase、量測槽與報告座標 | 日常維護主入口 |
| `DPO4054_POL_set\*.set` | 定義示波器通道、Probe、Scale、Trigger、Measurement | 儀器條件變更時 |
| Power Validation Report `.xlsm` | 提供 `VR SPEC` 輸入並接收圖片與數值 | 每個專案建立副本後使用 |
| `eScope_viewer.py` | 提供 `SCPIController`；另含示波器遠端 GUI | 引擎目前只匯入 Controller |
| `AutoTest_Config_Manual.md` | 舊版欄位說明 | 可參考，但以本 SOP 與程式現況為準 |
| 舊版 DC Test Gen 範本活頁簿 | 舊版 Excel/VBA 控制台、測項設定、ScopeSet 與回填對映的參考基準 | 不由 SemiAutoRunner 執行；不得以此覆蓋現行 Config 或報告 |
| `generate_config.py`、`add_*`、`fix_target.py`、`update_paths.py`、`explore_*` | 一次性產生、修改或探索用工具 | 不列入正式執行流程 |
| 根目錄 `ConfigGen.txt`、`LoBoSGen.txt` | 舊式 LoBo VBA 原始碼參考 | 不由 SemiAutoRunner 執行 |

一次性 Python 工具會用 pandas 重寫整本 `AutoTest_Config.xlsx`。這可能移除既有格式、驗證或其他原生 Excel 特性，因此不可直接對正式設定檔執行；必須先備份並在副本驗證差異。

## 4. 操作原理

### 4.1 正常資料流

```text
雙擊測項 .bat
  -> 切換至 SemiAutoRunner 目錄
  -> .venv Python 執行 general_auto_engine.py "Test_Item"
  -> 讀取 AutoTest_Config.xlsx 三張表
  -> 綁定已開啟的同名報告，或依 REPORT_PATH 開啟報告
  -> 從報告 VR SPEC 讀取輸入並計算負載提示
  -> 連線 192.168.1.10 的 DPO4054
  -> 逐行推送本機 .set 內的 SCPI 指令
  -> 每個 Phase 提示人工調整電子負載
  -> 人工確認波形穩定後按 Enter
  -> 查詢指定 MEAS 統計值或 Cursor Delta
  -> 優先用 HTTP /image.png 取圖，失敗才用 SCPI 檔案法
  -> 寫入報告儲存格並替換錨點附近舊圖
  -> 全部 Phase 結束後儲存報告
```

### 4.2 「半自動」的人工與自動邊界

| 人工負責 | 程式負責 |
| --- | --- |
| 確認 DUT、探棒、電子負載及電源接線安全 | 讀取報告與設定檔 |
| 依核准 Test Plan 設定輸入條件 | 推送 `.set` 至示波器 |
| 依畫面提示設定電子負載 | 顯示由 `VR SPEC` 算出的負載提示 |
| 確認通道、單位、波形及量測統計有效 | 查詢量測值、取得截圖 |
| 決定是否可以按 Enter | 將資料與圖片寫入指定報告位置 |
| 驗收報告，不合理時拒收並重測 | 全部流程結束時儲存報告 |

### 4.3 Excel 綁定規則

引擎先用「檔名」尋找已開啟的 Excel 活頁簿。若找到同名活頁簿，就直接綁定，不再核對完整路徑；若找不到，才用程式內的 `REPORT_PATH` 開啟。因此：

- 同一時間只能開啟一份相同檔名的報告。
- 執行前必須核對 Excel 標題列及完整路徑。
- 若更換專案報告，只改檔名而不改 `REPORT_PATH` 仍可能寫入舊檔。
- 引擎儲存後不會自動關閉 Excel。

### 4.4 示波器連線與設定載入

- VISA 位址格式：`TCPIP::192.168.1.10::INSTR`。
- 本機 `Setup_File` 存在時：逐行傳送非空白、非 `#` 開頭的 SCPI。
- 本機路徑不存在時：程式會把該字串當成示波器內部路徑，執行 `RECAll:SETUp`。
- 單行 SCPI 傳送錯誤目前會被忽略；畫面即使顯示「推送完畢」，仍須人工核對通道及 Measurement。
- 每次載入後程式會強制開啟 CH1，可能改變原 `.set` 的通道顯示狀態，但不會重設 Measurement Type 或 Source。

### 4.5 擷取、換算與回填

| `Meas_Types` 寫法 | 實際讀取 |
| --- | --- |
| `MAX`、`MAXimum` | 若未指定槽號，按清單位置使用 MEAS1、MEAS2…，查 `MAXimum?` |
| `MIN`、`MINImum` | 查指定 MEAS 的 `MINImum?` |
| `VAL` | 查指定 MEAS 的 `VALue?` |
| `MEAS3:MAX` | 明確查 MEAS3 的歷史 Maximum |
| 含 `FREQ` | 查 `VALue?`，先由 Hz 除以 1,000,000 成 MHz，再於寫入時乘 1,000，最後以 kHz 字串保存 |
| 含 `CURSOR` | 查垂直游標差 `CURSor:VBArs:DELTa?`，不使用 MEAS 槽 |
| 含 `TIME`、`WIDTH`、`JITTER` 或 `CURSOR` | 以秒乘 `1E9`，寫成 ns 字串 |
| 其他 | 一律乘 1,000 後以純數值寫入 |

「其他一律乘 1,000」代表程式沒有真正的工程單位模型。新增電流、電壓或其他物理量前，必須以已知輸入驗證示波器回傳單位與報告欄位單位，不能只看關鍵字名稱。

## 5. 執行前條件

### 5.1 人員與安全

- 執行者須具備 DUT、電源、電子負載與示波器操作資格。
- 電壓、電流、OCP、SCP 與切換節點測試條件必須以核准 Test Plan 為準。
- 在變更探棒、短路條件或高低側量測前，先將 DUT 置於安全狀態。
- 不確定接地、差動探棒額定值、電流探棒方向或負載能力時，不得開始測試。
- SOP 只描述軟體流程，不取代實驗室 ESD、電氣安全及設備額定規範。

### 5.2 必要軟硬體

| 類別 | 需求 |
| --- | --- |
| 作業系統 | Windows，且可執行 `.bat`、Excel COM |
| Excel | 桌面版 Microsoft Excel，可開啟及儲存 `.xlsm` |
| Python | 現有環境為 CPython 3.8.20 |
| Python 套件 | pandas 2.0.3、xlwings 0.31.10、requests 2.32.4、Pillow 10.4.0、PyVISA 1.14.1、pywin32 311、openpyxl 3.1.5 |
| VISA | NI-VISA、TekVISA 或其他可讓 `pyvisa.ResourceManager()` 找到儀器的 backend |
| 示波器 | DPO4054，LAN IP 預設為 `192.168.1.10` |
| 其他設備 | 依 Test Plan 的電源、電子負載、DUT、探棒與治具 |

### 5.3 每次測試的備份規則

1. 從乾淨 Sample 或上次核准版本複製一份專案報告。
2. 檔名包含專案、Rail、階段與日期，避免只用 `Sample`。
3. 執行前再複製一份 `*_before_run.xlsm` 備份。
4. 關閉不相關的同名 Excel 檔。
5. 確認 `general_auto_engine.py` 的 `REPORT_PATH` 指向本次報告。

## 6. 既有環境保護與健檢 SOP

### 6.1 Golden Environment 凍結原則

目前 `.bat` 已由操作者確認可正常執行，因此現有 `.venv` 視為 Golden Environment：

- 不刪除、不重建、不搬移 `.venv`。
- 不執行 `pip install`、`uv sync`、套件升級或 Python 升版。
- 不使用一次性工具重寫正式 Config 或報告。
- 需要驗證時優先執行唯讀檢查，不變更套件或環境變數。
- 若未來真的失效，先完整複製專案與環境，再於隔離副本研究修復；正式環境需保留原狀以便回復。

若未來必須升版，應建立另一個平行測試環境，完成所有測項回歸後再決定是否切換，不得在現有環境直接升級。

### 6.2 Python 唯讀健檢（僅在需要時）

```powershell
.venv\Scripts\python.exe --version
.venv\Scripts\python.exe -c "import pandas, xlwings, requests, pyvisa, PIL, win32com.client; print('Python dependencies OK')"
```

合格標準：兩個命令皆以 Exit Code 0 結束，且沒有 ImportError。

### 6.3 網路、HTTP 與 VISA 健檢

```powershell
Test-Connection 192.168.1.10 -Count 2
Invoke-WebRequest http://192.168.1.10/image.png -OutFile "$env:TEMP\dpo4054_preflight.png"
.venv\Scripts\python.exe -c "import pyvisa; rm=pyvisa.ResourceManager(); print(rm.list_resources()); rm.close()"
```

合格標準：

- Ping 或替代網路檢查可達。
- `image.png` 可下載且開啟後為目前示波器畫面。
- VISA 清單包含目標 TCPIP 儀器。
- 在開始正式測試前，再執行 `*IDN?` 確認回應為預期的 DPO4054。

## 7. 報告檔與 `VR SPEC` 填寫 SOP

### 7.1 輸入來源

引擎固定從報告工作表 `VR SPEC` 取值。典型電源驗證報告（以 1.8V 軌為例）的核心測試輸入為：

| 儲存格 | 目前快取值 | 用途 |
| --- | ---: | --- |
| `I36` | 2.6 | Imax / Full load |
| `H36` | 1.3 | Mid load / TDC，公式為 I36 ÷ 2 |
| `I37` | 2.6 | Dynamic/OCP 公式基準 |
| `G40`、`I40` | 0.26、1.3 | Dynamic Phase 1 範圍 |
| `G41`、`I41` | 1.3、2.6 | Dynamic Phase 2 範圍 |

### 7.2 填寫規則

1. 先依 Test Plan 完成專案基本資料及 `VR SPEC`。
2. 確認 Imax、TDC、Dynamic 範圍等欄位是數值或可正常計算的公式。
3. 在 Excel 手動重新計算，確認沒有 `#REF!`、`#VALUE!`、`#N/A`。
4. 把每個由程式使用的值寫入測試紀錄或 Run Sheet，作為提示值的人工對照。
5. 空白不可代表零。程式會把透過 `Spec` 讀到的空白轉為 `0.0`，所以空白輸入會偽裝成有效零值。
6. 儲存報告後再啟動 `.bat`，避免引擎讀到尚未儲存或尚未計算的狀態。

## 8. `AutoTest_Config.xlsx` 填寫 SOP

### 8.1 三表分工

```text
Item_Config     = 做哪一個測項、寫哪張報表、載入哪個 .set
Formula_Macro   = 從 VR SPEC 讀什麼、算出哪些人工操作提示
Layout_Mapping  = 每個 Phase 何時擷取、讀哪個 MEAS、寫到哪裡
```

三張表以 `Test_Item` 為主鍵；`Formula_Macro.Var_Name` 再與 `Layout_Mapping.Load_Variable` 關聯。所有比對皆為逐字比對，大小寫、空白、底線及標點都必須一致。

### 8.2 `Item_Config`

| 欄位 | 必填 | 規則 |
| --- | --- | --- |
| `Test_Item` | 是 | 全系統唯一 ID；須與 `.bat` 參數、Formula、Layout 完全一致 |
| `Enable` | 是 | 僅接受大寫 `YES` 或 `NO` |
| `Report_Sheet` | 是 | 必須是目標 `.xlsm` 中已存在的工作表名稱 |
| `Setup_File` | 是 | 建議使用已確認存在的本機絕對路徑 |
| `Vout_Cell` | 否 | 目前 `general_auto_engine.py` 未使用 |

重要：透過現有 `.bat` 啟動時會傳入 `specific_test`，程式只依 `Test_Item` 篩選，**不檢查 `Enable`**。因此把某測項改成 `NO` 並不能阻止對應 `.bat` 執行。要停用測項，應同時移除或明確封存該 `.bat`，並修正程式的停用邏輯。

### 8.3 `Formula_Macro`

| 欄位 | 必填 | 規則 |
| --- | --- | --- |
| `Test_Item` | 是 | 必須存在於 `Item_Config` |
| `Var_Name` | 是 | 同一測項內唯一；Layout 以此名稱取值 |
| `Spec_Cell` | 條件式 | 使用 `Spec['I36']` 時，同測項至少有一列須把 `I36` 載入；不用時可留白 |
| `Python_Eval_Formula` | 是 | 合法且受信任的 Python 表達式 |

可用物件：

- `Spec`：預先讀入的 `{儲存格地址: 值}` 字典。
- `ws`：報告的 `VR SPEC` 工作表，可用 `ws.range('G40').value` 讀值。

範例：

```text
Full_load  | I36 | float(Spec['I36'])
Mid_load   | H36 | float(Spec['H36'])
Light_load |     | float(Spec['I36']) * 0.1
load_cond  |     | str(ws.range('G40').value) + ' to ' + str(ws.range('I40').value)
```

注意事項：

- `eval()` 可以執行 Python 程式碼。Config 只能由受信任維護者修改，不可接受外部來源公式。
- 公式錯誤會被轉成 `0.0`，流程不會自動中止。
- 重複 `Var_Name` 時，後一列會覆蓋前一列。
- 數字結果會四捨五入到小數點後 2 位；字串不處理。

### 8.4 `Layout_Mapping`

| 欄位 | 必填 | 規則 |
| --- | --- | --- |
| `Test_Item` | 是 | 必須存在於 `Item_Config` |
| `Phase_Step` | 是 | 同一測項從 1 開始連續遞增，不重複 |
| `Prompt_Name` | 是 | 明確寫出負載或操作條件，不使用亂碼 |
| `Load_Variable` | 是 | 必須存在於同測項的 `Formula_Macro.Var_Name` |
| `Picture_Anchor` | 是 | Excel 範圍，例如 `B28:K49` |
| `Vpp_Data_Cell` | 是 | 一或多個儲存格；多個用半形逗號分隔 |
| `Meas_Types` | 是 | 與資料格數量、順序一對一；建議都明確指定 `MEASn:` |

範例：

```text
Vpp_Data_Cell = K61, K62
Meas_Types    = MEAS1:MAX, MEAS2:MIN
```

第一筆結果寫 K61，第二筆結果寫 K62。兩側數量不一致時，程式不會主動報錯，可能只寫入部分資料。

### 8.5 Config 修改完成檢查

- [ ] `Test_Item` 在三張表完全一致。
- [ ] 同一測項的 `Var_Name` 無重複。
- [ ] 每個 `Load_Variable` 都能找到同測項變數。
- [ ] Phase 從 1 連續遞增。
- [ ] Report Sheet、圖片範圍及資料格均存在。
- [ ] `Vpp_Data_Cell` 與 `Meas_Types` 數量及順序一致。
- [ ] 每個 `MEASn` 的 Type、Source 與報告標籤一致。
- [ ] 每個非頻率、非時間量測都確認過 ×1,000 後的工程單位。
- [ ] Config 中的本機 `.set` 路徑全部存在。
- [ ] 在報告副本完成一個低風險 Phase 的端到端測試。

## 9. `.set` 示波器設定維護 SOP

### 9.1 每個 `.set` 必須同時核對的六類設定

1. 通道：啟用狀態及 CH1–CH4 Label。
2. 探棒：Probe Gain、衰減比、電流探棒方向。
3. 垂直：Scale、Offset、Coupling、Bandwidth。
4. 水平：Time/Div、Record Length、取樣模式。
5. Trigger：Source、Slope、Level、Mode。
6. Measurement：每個 MEAS 的 Type、Source1/Source2、統計顯示。

### 9.2 儲存與變更規則

1. 從已驗證 `.set` 複製新檔，不覆寫唯一的 Golden 檔。
2. 檔名需能對應 `Test_Item` 與邊緣條件，例如 `5.VDS_HS_Rise.set`。
3. 在示波器本機載入，確認畫面與 Test Plan 一致。
4. 以已知信號確認每個 MEAS 槽回傳的物理量與單位。
5. 匯出後以純文字檢查 Channel Label、Measurement Type/Source、Trigger 及 Scale。
6. 更新 `Item_Config.Setup_File`。
7. 在報告副本執行一個 Phase，核對畫面、查詢值、換算值與目標儲存格。

### 9.3 現有 `.set` 快速對照

| 測項 | 主要啟用通道 | 核心 Measurement |
| --- | --- | --- |
| Ripple | CH1 Vout、CH2 Iout | MEAS3 = CH1 PK2PK；Config 讀 MEAS3 Maximum |
| Dynamic | CH1 Vout、CH2 Iout | MEAS1 = CH1 Maximum；MEAS2 = CH1 Minimum |
| FSW | CH2 Iout、CH4 FSW | MEAS1 = CH4 Frequency |
| Switching HS | CH2 Iout、CH3 Vin | MEAS1/2 = MATH Maximum/Minimum |
| Switching LS | CH2 Iout、CH4 L/S | MEAS1/2 = CH4 Maximum/Minimum |
| Jitter | CH2 Iout、CH4 Phase | MEAS1 = CH4 Negative Width；另使用 Cursor Delta |
| OCP/SCP | CH1 Vout、CH2 Iout、CH3 PG、CH4 Phase Node | CH2 最大值在現有 `.set` 的 MEAS3 |
| Turn On/Off | CH1 Vout、CH2 EN、CH3 Vin、CH4 PG | Rise/Fall 及 Delay |

## 10. 日常測試執行 SOP

### 10.1 開始前 Gate

- [ ] 已閱讀第 2 節維護確認事項；若與本次測項相關，已保存 Console、示波器畫面與報告證據。
- [ ] 已建立報告 `before_run` 備份。
- [ ] Excel 中只開啟一份正確的同名報告。
- [ ] `VR SPEC` 相關數值與人工 Run Sheet 一致。
- [ ] `CONFIG_PATH`、`REPORT_PATH`、`Setup_File`、`SCOPE_IP` 正確。
- [ ] DUT、探棒、電子負載及供電接線通過安全檢查。
- [ ] HTTP 截圖與 VISA `*IDN?` 均成功。
- [ ] 示波器日期、時間及校正狀態符合實驗室規範。

### 10.2 執行

1. 儲存報告，但保持正確報告開啟。
2. 雙擊本次測項 `.bat`。
3. 核對 Console 顯示的設定檔、報告檔名及示波器 ID。
4. 若出現「示波器連線失敗／模擬展示模式」，立即按 `Ctrl+C` 中止，不可按 Enter。
5. 核對 Console 顯示的每個計算變數；`0.0`、`None`、亂碼或不合理範圍都必須中止。
6. 在示波器確認通道、Probe、Trigger、Measurement 及波形均正確。
7. 依 Phase 提示設定電子負載或 Test Plan 條件。
8. 等待熱穩定、波形穩定或 Test Plan 規定時間。
9. 對照示波器畫面與報告欄位，確認 Measurement 槽及單位。
10. 按 Enter 擷取一次。
11. 立即到報告檢查本 Phase 的數值與圖片；不合理時停止後續 Phase。
12. 重複至全部 Phase 完成。
13. 看到「報告儲存成功」後，仍需人工檢查檔案時間戳及重新開啟結果。

### 10.3 中止規則

任一情況發生時立即中止：

- 儀器離線、VISA 逾時、HTTP 截圖失敗且 fallback 不可靠。
- 示波器回傳 `9.91E+37`、空值、0.0 或明顯不合理值。
- 提示負載與 Run Sheet 不一致。
- Measurement Type/Source 與 Config 或報告標籤不一致。
- 截圖不是目前 Phase、圖片空白或貼錯位置。
- DUT、探棒、治具或負載出現安全異常。
- Excel 寫入錯誤、檔案鎖定、寫入錯工作表或錯報告。

中止後先使硬體回到安全狀態，再處理軟體。若程式尚未儲存，關閉報告時選擇不儲存；若已儲存，從 `before_run` 備份復原後再重測。

## 11. 現有測項執行與回填對照

| `.bat` | Phase | 負載或條件來源 | 量測與回填 | 目前狀態 |
| --- | --- | --- | --- | --- |
| `啟動_2.Ripple_測試.bat` | 3 | 0.1×Imax、TDC、Imax | MEAS3 PK2PK 的 Maximum；K50、K76、K102；圖片 B28/B54/B80 區塊 | 2026-09-14 下午，操作者確認完成 |
| `啟動_3.Dynamic_Load_測試.bat` | 2 | G40→I40、G41→I41 | CH1 Maximum/Minimum；K61/K62、K89/K90 | 2026-09-14 下午，操作者確認完成 |
| `啟動_4.FSW_測試.bat` | 3 | 0.1×Imax、TDC、Imax | CH4 Frequency；J60、J86、J112，kHz 字串 | 2026-09-14 下午，操作者確認完成；報告值依第 14 節驗收 |
| `啟動_5.Switching_waveforms_測試.bat` | 4 個子測項，各 3 Phase | 0.1×Imax、TDC、Imax | HS 寫 K 欄、LS 寫 W 欄的 Max/Min；每個子測項結束即儲存 | 2026-09-14 下午，操作者確認完成；批次仍不會在子測項失敗時自動停止 |
| `啟動_6.Jitter_測試.bat` | 3 | 0.1×Imax、TDC、Imax | Cursor Delta 與 Pulse Width；J60/J61、J86/J87、J112/J113 | 2026-09-14 下午，操作者確認完成；量測槽列為維護確認事項 |
| `啟動_7.OCP_測試.bat` | 2 | Quick/Slow current rise | 寫 K70、K98 | 2026-09-14 下午，操作者確認完成；變數鍵與量測槽列為維護確認事項 |
| 無 `.bat` | SCP After/Before | 尚未定義 | 尚未定義 Layout | 未實作 |
| 無 `.bat` | Turn On/Off | 尚未定義 | 尚未定義 Layout | 未實作 |

Switching 批次會依序啟動四個獨立 Python 程序：HS Rise、HS Fall、LS Fall、LS Rise。批次檔沒有檢查 `ERRORLEVEL`，前一程序失敗時仍會繼續下一程序，最後也會印出 Completed。完成字樣不能作為測試成功證據。

### 11.1 2026-09-14 實機執行紀錄

| 欄位 | 紀錄 |
| --- | --- |
| 日期與時段 | 2026-09-14 下午 |
| 執行者 | 操作者本人 |
| 範圍 | 當日要求的全部現有測試 `.bat` |
| 操作結果 | 操作者確認全部完成 |
| 環境結果 | 現有 `.venv` 與雙擊啟動流程正常 |
| 證據邊界 | 本次 SOP 更新未重新操作實機，也未獨立審核當日報告數值與波形 |

## 12. 維護變更 SOP

### 12.1 依變更類型判斷修改範圍

| 變更需求 | 必改 | 必驗證 |
| --- | --- | --- |
| 換專案報告 | `REPORT_PATH`、報告 `VR SPEC` | 工作表名稱、輸入值、寫入位置、備份 |
| 搬資料夾 | 原則上不要搬動 Golden Environment；若必須搬移，先建立完整副本並更新 `CONFIG_PATH`、`REPORT_PATH`、全部本機 `Setup_File` | 副本 `.bat` 可啟動、所有路徑存在、至少一個端到端 Phase；原環境保持不變 |
| 換示波器/IP | `SCOPE_IP` | Ping、HTTP image、VISA、`*IDN?`、截圖 fallback |
| 改負載條件 | `Formula_Macro` 或報告 `VR SPEC` | 人工計算對照、空白/零值、邊界值 |
| 改量測槽 | `.set` 與 `Layout_Mapping.Meas_Types` | Type、Source、統計欄、單位、報告標籤 |
| 改回填位置 | `Picture_Anchor`、`Vpp_Data_Cell` | 目標無公式被覆蓋、圖片尺寸、舊圖替換 |
| 新增測項 | Config 三表、`.set`、新 `.bat` | 第 13 節完整流程 |
| 改引擎 | Python 與可重現依賴 | 全測項報告副本回歸、錯誤路徑、離線 fail-closed |

### 12.2 路徑維護

目前硬編碼位置：

- `general_auto_engine.py`：`CONFIG_PATH`、`REPORT_PATH`、`SCOPE_IP`。
- `generate_config.py`、`update_paths.py`、`fix_target.py`、`explore_*`：各自有舊絕對路徑。
- `AutoTest_Config.xlsx`：每列 `Setup_File`。
- `.venv\pyvenv.cfg`：建立環境時的 Python home。

不要用舊 `update_paths.py` 當通用搬移工具。它只更新 Ripple 與 Dynamic，且指向 `SemiAutoRunner\Oscilloscope_Test_set_files`，與目前正式 Config 使用的根目錄 `DPO4054_POL_set` 不一致。

### 12.3 版本與變更紀錄

每次正式變更至少保存：

- 變更單號、日期、作者、原因。
- Config 前後版本。
- 受影響 `.set` 的雜湊或版本。
- 測試用報告副本與前後截圖。
- 實機、通道、Probe、DUT、Rail、負載條件。
- 回歸項目、預期、實際與結論。
- 已知限制及回復方式。

目前 Git 儲存庫尚無任何 commit，所有檔案均未追蹤。建立第一個基準版本前，應先排除 Excel 暫存檔 `~$*`、`.venv`、輸出圖片、cache 與測試報告，再由維護者審核要納入版本控制的原始碼、Config 範本與 `.set`。

## 13. 新增測項 SOP

1. 定義 Test Plan：目的、輸入條件、負載、通道、Probe、Trigger、Measurement、單位、允收標準。
2. 在報告副本確認目標工作表、圖片框及資料儲存格；不得覆蓋公式或合併儲存格主體之外的位置。
3. 從 Golden `.set` 複製新檔，在示波器調整並以已知信號驗證。
4. 在 `Item_Config` 新增唯一 `Test_Item`。
5. 在 `Formula_Macro` 新增唯一變數，先用人工計算核對結果。
6. 在 `Layout_Mapping` 依 Phase 新增行，明確使用 `MEASn:`。
7. 建立 `.bat`，其引數必須與 `Test_Item` 完全一致。
8. 在 Config 上執行第 8.5 節一致性檢查。
9. 使用報告副本，只跑 Phase 1；核對 Console、示波器、原始查詢、換算值、圖片及報告座標。
10. 再跑其餘 Phase 與邊界條件。
11. 驗證儀器離線、查詢失敗、Excel 鎖定、目標 Sheet 不存在時會安全失敗，不寫入正式結果。
12. 完成同類既有測項的回歸後才發版。

批次檔範本：

```batch
@echo off
title New Semi-Auto Test
cd /d "%~dp0"

.venv\Scripts\python.exe general_auto_engine.py "Exact_Test_Item"
if errorlevel 1 (
  echo Test failed. Report was not accepted.
  pause
  exit /b 1
)

echo Test completed. Perform report acceptance checks.
pause
```

正式版本還應讓 Python 在連線、查詢、截圖或回填失敗時回傳非零 Exit Code，否則批次檔無法可靠攔截失敗。

## 14. 輸出驗收 SOP

### 14.1 每個 Phase 驗收

- [ ] 圖片是本次 DUT、本次負載與本次 Phase。
- [ ] 圖片中的 Channel Label、Scale、Time/Div、Trigger 及 Measurement 可辨識。
- [ ] 報告內圖片完整、比例正確、未蓋住文字且舊圖已被取代。
- [ ] 儲存格標籤與 `MEASn` Source/Type 一致。
- [ ] 數值與示波器畫面一致，並套用正確單位轉換。
- [ ] 無 `9.91E+37`、`0.0` fallback、空白、`None` 或明顯異常量級。
- [ ] 負載提示與人工 Run Sheet 一致。

### 14.2 全測項驗收

- [ ] 所有預期 Phase 都有圖片及數值，沒有缺頁或重複圖。
- [ ] 報告沒有 `#REF!`、`#VALUE!`、`#DIV/0!`。
- [ ] 測試結果與規格的 Pass/Fail 公式正確更新。
- [ ] 檔案可儲存、關閉、重新開啟，巨集與圖片仍存在。
- [ ] 最終報告與 `before_run` 差異只包含預期儲存格、圖片與計算結果。
- [ ] 測試紀錄包含環境、儀器、DUT、Config、`.set` 與報告版本。

## 15. 故障排除

| 症狀 | 最可能原因 | 處理順序 |
| --- | --- | --- |
| `.bat` 未來突然無法啟動 | 路徑、權限、防毒軟體或 `.venv` 狀態改變 | 先保留現場並做唯讀檢查；不要直接重建或安裝套件，在完整副本排查 |
| 顯示找不到測項 | `.bat` 引數與 `Test_Item` 不一致 | 比對大小寫、空白、底線、標點 |
| 設為 `NO` 仍執行 | `.bat` 傳入 specific test，繞過 Enable | 中止並修正批次/引擎；不要只改 Enable |
| 負載提示為 0.0 | Spec 空白、公式錯誤、變數名稱不一致 | 先查 Console 錯誤，再查 Formula 與 Layout 鍵值 |
| 進入模擬模式 | VISA、IP、驅動或網路失敗 | 立即中止，不按 Enter；依第 6.3 節排查 |
| 頻率為極大值 | 示波器無有效週期或回傳 sentinel | 檢查信號、Trigger、MEAS1；拒收本次資料 |
| Jitter Pulse Width 極大 | `VAL_WIDTH` 讀到 MEAS2 RMS | 明確指定正確 MEAS 槽並實機核對 |
| 數值大或小 1,000 倍 | `Meas_Types` 關鍵字或工程單位錯 | 核對原始 SCPI 值、程式換算及報告單位 |
| 圖片沒貼上 | HTTP 失敗、示波器 E: fallback 不可用、本機圖不存在、工作表保護 | 先看 Console，再確認本機 PNG 及工作表狀態 |
| 貼到錯報告 | 同名活頁簿綁定或 `REPORT_PATH` 錯 | 立即停止；核對完整路徑並從備份復原 |
| 儲存失敗 | 檔案權限、鎖定、唯讀或 Excel 多執行個體 | 另存救援副本，關閉重複 Excel，勿覆寫原檔 |
| `.set` 顯示成功但畫面不對 | 單行 SCPI 錯誤被忽略 | 在示波器逐項核對通道、Trigger、Measurement |
| Switching 顯示 Completed 但缺資料 | 子程序失敗後批次仍繼續 | 逐段查 Console、圖片及報告，不接受 Completed 字樣 |

## 16. 維護完成定義

一次變更只有同時符合以下條件才算完成：

1. 結構正確：Config 三表、`.bat`、`.set`、報告工作表與座標一致。
2. 環境未受污染：Golden Environment 未被重建、升級或同步套件，現況版本另有紀錄。
3. 儀器可驗證：HTTP、VISA、`*IDN?`、Measurement Source/Type 均有證據。
4. 結果可信：沒有 dummy、sentinel、靜默 0.0 或錯誤單位。
5. 失敗安全：離線、查詢、截圖、Excel 或儲存失敗時不會產生正式結果。
6. 回歸完成：受影響測項及相鄰測項在報告副本通過。
7. 可回復：保留變更前 Config、`.set`、程式與報告備份。
8. 可追溯：變更原因、版本、測試人、日期與結論均有紀錄。

## 17. GUI 操作與維護規格

### 17.1 舊版活頁簿可沿用的操作模型

舊版測試控制活頁簿的 `VR TestPlan Control Sheet` 同時提供逐項測試按鈕與 YES/NO 批次選擇；`Item Config`、`ScopeSet_Mapping`、`C_Data_Mapping` 則分別承擔測試步驟、示波器設定與報告回填對映。新 GUI 應保留這四個概念，但不得把所有 Excel 欄位直接呈現給日常操作者。

| 舊版設計 | 新 GUI 決策 | 理由 |
| --- | --- | --- |
| Test Plan 檔名、路徑、版本與 Client Code 放在控制首頁 | 放入唯讀的「本次執行內容」列；更換報告必須重新通過預檢 | 避免寫錯專案或錯報告 |
| 每個測項各有一個巨集按鈕 | 改成測項清單加一個隨狀態變化的主要按鈕 | 降低按錯巨集及重複入口 |
| YES/NO 欄位搭配 `AutoTest` | 明確提供「單項執行」及「今日批次」兩種模式 | 同時支援除錯重測與排定工作 |
| `GPIB SCAN` | 轉成執行前預檢；顯示儀器型號、位址與驗證時間 | 連線成功不等於目標正確 |
| `KEEP/DELETE`、ScopeSet、Data Mapping 在設定表直接編輯 | 移到維護模式，預設唯讀；修改前顯示影響範圍並另存草稿 | 防止測試中誤改正式設定 |
| Protect/Unprotect 與密碼欄位在首頁 | 只供維護者角色使用，不在操作者首頁顯示，也不得儲存明碼密碼 | 操作與權限管理分離 |
| `No Stopped Auto Test` | 不沿用；批次預設遇錯即停，保留已完成證據後由人員決定重試或續跑 | 避免錯誤一路傳到正式報告 |
| 手動按 `Version Record Update` | 每次開始、回填、重試、中止及設定變更都自動寫稽核紀錄 | 降低漏記版本與責任資訊 |

### 17.2 日常操作者畫面

日常畫面只負責「選、看、做、驗」四件事，資訊不得重複：

1. 選：選擇報告，以及單項執行或今日批次；批次模式用核取方塊排定測項。
2. 看：中央區使用 eScope 即時投影，固定顯示 DUT/報告完整路徑、Test Item、ScopeSet、回填位置、DPO4054 身分、畫面時間及新鮮度。
3. 做：一次只顯示目前 Phase 的負載值、接線/探棒提示與一個主要動作。
4. 驗：同頁顯示最新示波器畫面、原始量測值、換算後值、圖片檔與 Excel 回填座標。

主要按鈕依狀態只可為下列之一：`開始預檢`、`開始測項`、`擷取並回填`、`重試本 Phase`、`完成驗收`。`安全中止` 必須獨立放置，且不得與主要按鈕共用顏色或位置。

### 17.3 維護者畫面

維護模式分成三個互不重疊的頁籤，名稱與現行 Config 責任一致：

1. 測項：`Item_Config` 的 Enable、Test Item、報告工作表及 Setup File。
2. 計算：`Formula_Macro` 的輸入儲存格、變數名稱、公式及計算預覽。
3. 版面：`Layout_Mapping` 的 Phase、負載、Prompt、量測類型、圖片與數值座標。

每次編輯必須即時驗證 `Test_Item`、`Var_Name`、`.set` 路徑、`MEASn` 與報告座標；存檔前列出受影響的 `.bat`、Phase、`.set` 和回填位置。正式 Config 不可在測試執行中修改，也不得由 GUI 自動安裝套件、重建 `.venv` 或同步環境。

### 17.4 GUI 狀態與失敗規則

GUI 必須以狀態機控制流程：`未就緒 -> 預檢 -> 等待人工設定 -> 擷取 -> 回填驗證 -> 完成/中止`。任一時間只允許一個儀器操作；新畫面不得覆蓋剛完成命令後的證據畫面。

- VISA、HTTP、Excel、目標身分或路徑未通過時，主要按鈕保持停用。
- 進入模擬模式、量測回傳 sentinel/空白、圖片失敗或回填失敗時，不得顯示完成，也不得自動跳到下一 Phase。
- 批次失敗時保留已完成測項，失敗測項標記為「需處理」，其後測項保持「未執行」。
- 重試只重做目前 Phase，不得重寫其他已驗收資料。
- 關閉 GUI 時若儀器操作仍在進行，先要求安全中止；`Stop Live`、中斷測試與中斷儀器連線是三個不同動作。

#### eScope 投影整合邊界

GUI 的中央畫面由 `$tektronix-escope-control` 的 Observe 路徑提供。投影端只允許 `GET /image.png?N=<整數>`，不得接受任意 path、query string、`COMMAND=` 或原始 SCPI。HTTP 投影與 VISA 控制是兩條獨立通道，必須分別顯示狀態：

1. 投影狀態：目標 IP、最後畫面時間、畫面年齡、PNG 尺寸及 Live/停止/失聯。
2. 控制狀態：VISA 是否連線、`*IDN?` 是否符合指定 DPO4054、目前是否有操作進行中。
3. 證據狀態：命令前畫面、命令後畫面、語意驗證及復原狀態；HTTP 200、圖片位元數或 Viewer 的 `OK:` 不得當成控制成功證據。

執行緒與資源管理必須符合下列規則：Tkinter Widget 與 `ImageTk.PhotoImage` 只在主執行緒建立；HTTP、VISA 及 Pillow 處理放在背景執行緒；投影與控制操作序列化；舊 Live Frame 不得覆蓋命令後證據畫面；關閉視窗時停止輪詢並關閉 HTTP/VISA 資源。現行 Viewer 的有效 50 ms 最低更新間隔及 LANCZOS 顯示行為維持不變，除非另案明確重啟效能評估。

### 17.5 GUI 第一階段驗收條件

第一階段只建隔離原型，不改現有 `.bat`、`.venv`、正式 Config 或示波器。原型通過後才決定是否串接現行引擎。

- [ ] 單項及今日批次都能用 1024×768 投影 Fixture 及模擬資料走完狀態流程。
- [ ] Observe 測試產生的 URL 只能包含目標 Host、`/image.png` 與整數 `N`，不可產生 `COMMAND=`。
- [ ] 投影中斷只標記 Observe 異常，不得錯誤顯示 VISA 已斷線；反向亦同。
- [ ] 模擬舊 Frame 延遲到達時，不會覆蓋命令後證據畫面。
- [ ] 一般操作者看不到原始 SCPI、密碼、Protect/Unprotect 或 Config 儲存格。
- [ ] 畫面可追溯 `Test Item -> Setup File -> Phase -> MEAS/Cursor -> Excel 座標`。
- [ ] 未確認負載與波形前，無法擷取回填。
- [ ] 任一模擬錯誤都會 fail-closed，且不顯示「完成」。
- [ ] 安全中止、重試、續跑及關閉視窗的結果均明確且可恢復。
- [ ] Golden Environment 在驗收前後沒有任何套件或檔案變更。
