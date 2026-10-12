# Power-SemiAutoRunner 電源驗證半自動化測試系統

> **專為硬體/電源驗證工程師設計的高效量測與自動回填架構**  
> 串接示波器 SCPI 控制、參數動態計算、波形自動擷取與 Excel 測試報告回填，並預留 AI 代理人操作接口。

---

## 📌 專案簡介 (Overview)

在伺服器、主機板與電源模組（VR / POL）驗證過程中，工程師常需耗費大量時間進行繁瑣的重複操作：手動切換示波器設定、手動調整電子負載、手動儲存截圖、手動將量測數值與圖表複製貼上至 Excel 報告中。

**Power-SemiAutoRunner** 將此流程標準化、自動化：
1. **儀器配置自動化**：根據測項自動推送指定示波器設定檔（`.set`）。
2. **負載智慧提示**：自動根據報告中之規格定義（如 `VR SPEC`），動態計算輕載（10% Imax）、中載（TDC）、滿載（Imax）電流，並於終端清晰提示。
3. **單步波形確認**：波形穩定後按確認鍵，系統立即自動完成截圖。
4. **精準量測與單位轉換**：透過 SCPI 讀回示波器統計值（Peak-to-Peak、Max、Min、Frequency、Jitter、Delay 等），並依測項自動換算單位（mV、V、kHz、ns）。
5. **Excel 報告自動對位回填**：利用 `xlwings` 即時將波形圖自動貼入指定儲存格範圍，並將量測數值精準填入對應欄位，自動存檔。
6. **AI Agent 擴充就緒**：內建 `AIAgentBridge`，支援視覺快照（Base64/PIL）與遠端 SCPI 調度，便於未來接入多模態 AI 進行自主判讀。

---

## 🛠️ 支援測試項目 (Supported Test Items)

| 編號 | 測試項目 (Test Item) | 說明 | 量測類型 |
| :---: | :--- | :--- | :--- |
| **02** | **Ripple & Noise** | 紋波與雜訊測試（輕載、中載、滿載） | Peak-to-Peak (mV) |
| **03** | **Dynamic Load** | 動態負載瞬態響應測試（Phase 1 / Phase 2） | Vout Max / Min (V) |
| **04** | **FSW** | 切換頻率測試（輕載、中載、滿載） | Frequency (kHz) |
| **05** | **Switching Waveforms** | MOSFET VDS 開關波形測試（高側/低側 上升/下降沿） | Vds Max / Min (V) |
| **06** | **Jitter** | 開關節點 Jitter 與脈寬穩定度測試 | Jitter Cursor, Pulse Width (ns) |
| **07** | **OCP** | 過電流保護點測試 | Trigger Peak Current (A) |
| **08** | **SCP** | 短路保護測試（上電前短路 / 上電後短路） | Short Peak Current (A) |
| **10** | **Turn-on & off Sequence** | 開關機時序與延遲量測（上升/下降時間、EN-to-Vout 延遲） | Rise/Fall Time, Delay Time (ms) |

---

## 📐 系統架構 (Architecture)

系統採用嚴格的 **MECE 分層解耦設計**：

```
SemiAutoRunner
├── general_auto_engine.py      # 核心測試引擎 (排程控制、計算、擷取、回填)
├── scpi_controller.py          # 標準 VISA 示波器 SCPI 通訊模組
├── ai_agent_bridge.py          # AI 代理人操作橋接器 (視覺擷取、自主決策 API)
├── AutoTest_Config_Manual.md   # 設定檔完整填寫說明書
├── SemiAutoRunner_MECE_SOP.md  # 系統標準作業規範 (SOP)
├── VR_POL_AI_Autonomous_Testing_Proposal.md # 專題匯報講稿與 AI 自主量測演進提案
├── pyproject.toml              # 專案相依性配置
└── 啟動_*.bat                  # 各測項一鍵啟動批次檔
```

### 設定檔三層結構 (`AutoTest_Config.xlsx`)
- **`Item_Config`**：測項清單、啟用開關（`Enable`）、目標報告工作表（`Report_Sheet`）與示波器設定檔路徑。
- **`Formula_Macro`**：動態參數計算公式，支援動態讀取報告儲存格基準值（如 `Spec['I36'] * 0.1`）。
- **`Layout_Mapping`**：步驟排程、圖片回填錨點（`Picture_Anchor`）、數值目標儲存格（`Vpp_Data_Cell`）與量測模式（`Meas_Types`）。

---

## 🚀 快速開始 (Getting Started)

### 1. 建立 Python 虛擬環境

推薦使用現代高速套件管理器 [uv](https://github.com/astral-sh/uv)：

```bash
# 建立虛擬環境
uv venv .venv

# 啟動虛擬環境 (Windows PowerShell)
.venv\Scripts\Activate.ps1

# 安裝所需相依性
uv pip install xlwings pandas openpyxl pyvisa requests pillow
```

### 2. 環境與儀器連線需求
- **作業系統**：Windows 10 / 11（支援 Microsoft Excel 桌面版）。
- **示波器通訊**：支援 VISA LAN/TCPIP 通訊之示波器（如 Tektronix DPO/MSO 系列）。
- **預設 IP 位址**：`192.168.1.10`（可透過環境變數 `AUTOTEST_SCOPE_IP` 覆寫）。

### 3. 執行測試

可直接雙擊對應測項的批次檔：
- `啟動_2.Ripple_測試.bat`
- `啟動_3.Dynamic_Load_測試.bat`
- `啟動_4.FSW_測試.bat`
- `啟動_5.Switching_waveforms_測試.bat`
- `啟動_6.Jitter_測試.bat`
- `啟動_7.OCP_測試.bat`
- `啟動_8.SCP_測試.bat`
- `啟動_10.Turn-on_off_sequence_測試.bat`

或透過命令列指定測項執行：
```bash
.venv\Scripts\python.exe general_auto_engine.py "2.Ripple_and_Noise_POL"
```

---

## 🤖 AI 代理人介面 (AI Agent Integration)

本專案內建 `AIAgentBridge`，支援與多模態 AI Agent（如 Gemini 2.5 / 3.0 Vision）整合：

```python
from ai_agent_bridge import AIAgentBridge
from general_auto_engine import AutoTestEngine

engine = AutoTestEngine()
ai_bridge = AIAgentBridge(engine=engine)

# 1. 擷取當前示波器畫面 Base64 (供 LLM 視覺模型分析波形)
image_base64 = ai_bridge.get_screen_base64()

# 2. 根據 AI 判讀結果動態調整示波器設定
ai_bridge.send_scpi_command("CH1:SCALE 0.05")

# 3. 自動推進測試步驟
ai_bridge.confirm_step()
```

---

## 📖 說明文件 (Documentation)

- [VR/POL 測試專題匯報與 AI 自主量測演進提案 (VR_POL_AI_Autonomous_Testing_Proposal.md)](./VR_POL_AI_Autonomous_Testing_Proposal.md)：提供 10 分鐘簡報大綱、現場實戰逐字講稿、MECE 異常自癒架構與高階主管 Q&A 實戰攻防劇本。
- [AutoTest_Config 填寫說明書 (AutoTest_Config_Manual.md)](./AutoTest_Config_Manual.md)：詳細介紹每張工作表的填寫規範、公式寫法與除錯指南。
- [SemiAutoRunner MECE 作業標準 (SemiAutoRunner_MECE_SOP.md)](./SemiAutoRunner_MECE_SOP.md)：詳細定義執行環境、資料流、安全防護與回檢作業標準。

---

## ⚠️ 免責聲明 (Disclaimer)

本專案旨在輔助量測數據擷取與報告文書處理。操作人員進行高功率或具風險之電源測試（如 OCP、SCP）時，仍應嚴格遵守實驗室安全規範，手動確認接線與散熱條件，切勿依賴自動化程式替代硬體安全判斷。
