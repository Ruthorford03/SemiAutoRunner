import os
import sys
import time
import xlwings as xw
import pandas as pd

from scpi_controller import SCPIController

# ================= Configuration =================
# 分離後的配置檔案與實際要輸出的報告檔（支援相對路徑與環境變數覆寫）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.environ.get("AUTOTEST_CONFIG_PATH", os.path.join(BASE_DIR, "AutoTest_Config.xlsx"))
REPORT_PATH = os.environ.get("AUTOTEST_REPORT_PATH", os.path.join(BASE_DIR, "Power_Validation_Report_Sample.xlsm"))
SHEET_SPEC = "VR SPEC"

SCOPE_IP = os.environ.get("AUTOTEST_SCOPE_IP", "192.168.1.10") 
# =================================================

class AutoTestEngine:
    def __init__(self, config_path=None, report_path=None, scope_ip=None, log_callback=None):
        self.config_path = config_path or CONFIG_PATH
        self.report_path = report_path or REPORT_PATH
        self.scope_ip = scope_ip or SCOPE_IP
        self.log_callback = log_callback

        self.log(f"[*] 載入測試設定庫: {self.config_path}")
        self.df_item = pd.read_excel(self.config_path, sheet_name="Item_Config")
        self.df_formula = pd.read_excel(self.config_path, sheet_name="Formula_Macro")
        self.df_layout = pd.read_excel(self.config_path, sheet_name="Layout_Mapping")
        
        report_filename = os.path.basename(self.report_path)
        try:
            # 嘗試綁定已經在背景開啟的同名 Excel 檔案
            self.wb_report = xw.books(report_filename)
            self.app = self.wb_report.app
            self.log(f"[*] 已綁定開啟中的報告檔: {report_filename}")
        except Exception:
            # 如果沒找到，就開啟一個新的
            self.app = xw.App(visible=True, add_book=False)
            self.app.display_alerts = False
            self.log(f"[*] 開啟測試報告檔: {self.report_path}")
            self.wb_report = self.app.books.open(self.report_path)
            
        self.ws_spec = self.wb_report.sheets[SHEET_SPEC]
        
        self.scpi = SCPIController(self.scope_ip)
        ok, msg = self.scpi.connect()
        self.isConnected = ok
        if not ok:
            self.log(f"[!] 示波器連線失敗: {msg} (將進入模擬展示模式)")
        else:
            self.log(f"[*] 示波器連線成功: {msg}")

    def log(self, msg):
        """ 統一的日誌輸出介面，若有註冊回呼則通知 GUI，並同步 print """
        if self.log_callback:
            try:
                self.log_callback(msg)
            except Exception:
                pass
        print(msg)

    def evaluate_variables(self, test_item):
        """ 解析 Formula_Macro 中的變數，並代入 VR_SPEC 的基準去換算 """
        formulas = self.df_formula[self.df_formula["Test_Item"] == test_item]
        Spec = {} # Dictionary to store cell values matching Spec_Cell dynamically
        
        # 1. 搜集所需要的 Cell Value (例如 'I36', 'H36')
        for idx, row in formulas.iterrows():
            cell_addr = row["Spec_Cell"]
            if str(cell_addr) != "nan" and str(cell_addr) != "N/A":
                Spec[cell_addr] = self.ws_spec.range(cell_addr).value
                # 防呆
                if Spec[cell_addr] is None: Spec[cell_addr] = 0.0

        # 2. 透過 Python 的 eval 計算字串方程式 (e.g., ws.range('G40').value )
        calculated_vars = {}
        for idx, row in formulas.iterrows():
            var_name = row["Var_Name"]
            equation = row["Python_Eval_Formula"]
            try:
                # 這裡的 eval 會使用區域變數 Spec 以及 ws (直接操控 Excel API)
                res = eval(equation, {"Spec": Spec, "ws": self.ws_spec})
                
                # 如果是字串 (例如 "12.5 to 15.0") 就直接存，如果是數字就 round
                if isinstance(res, str):
                    calculated_vars[var_name] = res
                else:
                    calculated_vars[var_name] = round(res, 2)
            except Exception as e:
                print(f"[!] 變數 {var_name} 計算錯誤: {e}")
                calculated_vars[var_name] = 0.0
                
        return calculated_vars
        
    def control_oscilloscope(self, setup_file):
        if not self.isConnected: return
        print(f"\n[*] 開始初始化示波器 (載入設定檔: {setup_file})...")
        
        # 若為本機路徑與檔案，則直接讀取對示波器推送
        if os.path.exists(setup_file):
            print("    - 偵測到本機設定檔，開始按行推送給示波器 (可能需要幾秒鐘)...")
            with open(setup_file, 'r', encoding='utf-8', errors='ignore') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        try:
                            self.scpi.write(line)
                        except:
                            pass
            self.scpi.query('*OPC?')
            print("    - 本機設定檔推送完畢！")
        else:
            # 否則假定其為示波器內的相對路徑 (例如 E:/ 隨身碟)
            self.scpi.write(f'RECAll:SETUp "{setup_file}"')
            
        self.scpi.write('SELECT:CH1 ON')
        self.scpi.query('*OPC?')
        print("    - 示波器設定載入完成！")
        
    def capture_waveform(self, local_img_path, query_stats=["MAX"]):
        if not self.isConnected: 
            return [0.05] * len(query_stats), "dummy.png"
        
        # 移除對 MEAS:TYPe 與 SOUrce 的強制覆寫，全權交給使用者準備好的 .set 檔決定
        # 我們只負責在使用者按下 Enter 後，把數值撈回來
        time.sleep(0.5)
        
        # 動態讀取測量結果
        vpp_vals = []
        for i, stat_raw in enumerate(query_stats):
            stat_raw = stat_raw.strip().upper()
            
            # 支援精準指定量測項 (e.g., "MEAS3:MAX", "MEAS2:MIN")
            if ":" in stat_raw:
                meas_src, stat = stat_raw.split(":", 1)
                meas_idx = meas_src.upper().replace("MEAS", "").strip()
            else:
                meas_idx = str(i + 1)
                stat = stat_raw
                
            try:
                # 決定擷取統計資料框中的哪一種數值
                if "CURSOR" in stat_raw.upper():
                    val = float(self.scpi.query('CURSor:VBArs:DELTa?'))
                elif stat == "MIN" or "MIN" in stat:
                    val = float(self.scpi.query(f'MEASUrement:MEAS{meas_idx}:MINImum?'))
                elif "VAL" in stat:
                    val = float(self.scpi.query(f'MEASUrement:MEAS{meas_idx}:VALue?'))
                    if "FREQ" in stat:
                        val = val / 1000000.0
                else: # 預設抓取歷史波峰最大 MAX 值
                    val = float(self.scpi.query(f'MEASUrement:MEAS{meas_idx}:MAXimum?'))
                vpp_vals.append(val)
            except Exception as err:
                print(f"    ❌ [ERROR] 取回 MEAS{meas_idx} 統計結果失敗: {err}")
                vpp_vals.append(None)
            
            if vpp_vals[-1] is not None:
                print(f"    - 取回 MEAS{meas_idx} 統計結果 ({stat}) : {vpp_vals[-1]:.4f} (若為mV: {vpp_vals[-1] * 1000:.2f})")
            else:
                print(f"    - 取回 MEAS{meas_idx} 統計結果 ({stat}) : [無效/讀取失敗]")
        
        # 優先嘗試透過 HTTP 截取螢幕 (速度最快且不佔用示波器 I/O)
        import requests
        try:
            r = requests.get(f"http://{self.scpi.ip}/image.png", timeout=5)
            r.raise_for_status()
            with open(local_img_path, 'wb') as f:
                f.write(r.content)
            print(f"    - 成功透過 HTTP 取得截圖: {local_img_path}")
        except Exception as e:
            print(f"    - HTTP 截圖失敗 ({e})，改用 SCPI 備用方法...")
            scope_temp_file = "E:/TempRipple.png"
            self.scpi.write('SAVe:IMAGe:FILEFormat PNG')
            self.scpi.write(f'SAVe:IMAGe "{scope_temp_file}"')
            self.scpi.query('*OPC?')
            
            original_timeout = self.scpi.scope.timeout
            self.scpi.scope.timeout = 60000
            try:
                image_data = self.scpi.scope.query_binary_values(
                    f'FILESystem:READFile "{scope_temp_file}"', 
                    datatype='B', container=bytes, header_fmt='ieee'
                )
            except Exception:
                try:
                    self.scpi.write(f'FILESystem:READFile "{scope_temp_file}"')
                    image_data = self.scpi.scope.read_raw()
                except Exception:
                    image_data = b""
            finally:
                self.scpi.scope.timeout = original_timeout

            if image_data:
                with open(local_img_path, 'wb') as f:
                    f.write(image_data)
                
            self.scpi.write(f'FILESystem:DELEte "{scope_temp_file}"')
            self.scpi.query('*OPC?')
            
        return vpp_vals, local_img_path

    def write_to_report(self, sheet_name, vpp_vals, img_path, anchor_cell, vpp_cells, meas_types=None):
        ws = self.wb_report.sheets[sheet_name]
        
        # 判斷是否為 Dynamic Load 測項（量測 Vout Max/Min，以 V 為單位，不乘 1000）
        is_dynamic = "DYNAMIC" in sheet_name.upper()
        # 判斷是否為 Jitter 測項（量測 Jitter Cursor 與 Pulse Width，以 ns 為單位，乘以 1e9）
        is_jitter = "JITTER" in sheet_name.upper()

        # 支援多欄位填寫，例如 K61,K62
        for i, cell in enumerate(vpp_cells):
            cell = cell.strip()
            if i < len(vpp_vals) and cell:
                if vpp_vals[i] is None:
                    print(f"    ⚠️ [WARN] 儲存格 {cell} 讀值無效 (通訊中斷或逾時)，填入 'MEAS_ERR' 提示覆核，嚴禁回填偽數據！")
                    ws.range(cell).value = "MEAS_ERR"
                    continue
                if meas_types and i < len(meas_types):
                    m_type = meas_types[i].upper()
                    if "FREQ" in m_type:
                        # 頻率數值目前為 MHz，乘以 1000 轉換為 kHz 並以字串填入
                        khz_val = vpp_vals[i] * 1000
                        ws.range(cell).value = f"{khz_val:.1f} kHz"
                    elif is_jitter or "TIME" in m_type or "WIDTH" in m_type or "JITTER" in m_type or "CURSOR" in m_type:
                        # 時間/Jitter 數值預設為秒 (s)，乘以 1e9 轉換為 ns 並以純數值填入（報表格式會自動顯示 ns）
                        ns_val = round(vpp_vals[i] * 1e9, 1)
                        ws.range(cell).value = ns_val
                    elif is_dynamic or "VOLT" in m_type or "RAW" in m_type or "DIRECT" in m_type or "UNIT_V" in m_type:
                        # Dynamic Load 或明確指定電壓/原始值單位：直接填入原始數值 (V)，不乘 1000
                        ws.range(cell).value = vpp_vals[i]
                    else:
                        ws.range(cell).value = vpp_vals[i] * 1000
                else:
                    if is_jitter:
                        ws.range(cell).value = round(vpp_vals[i] * 1e9, 1)
                    elif is_dynamic:
                        ws.range(cell).value = vpp_vals[i]
                    else:
                        ws.range(cell).value = vpp_vals[i] * 1000
        
        if img_path == "dummy.png": return # 模擬模式不貼圖
        abs_img_path = os.path.abspath(img_path)
        anchor = ws.range(anchor_cell)
        
        if ws.api.ProtectContents:
            try: ws.api.Unprotect()
            except: pass

        for pic in ws.pictures:
            if abs(pic.left - anchor.left) < 50 and abs(pic.top - anchor.top) < 50:
                pic.delete()

        ws.pictures.add(abs_img_path, top=anchor.top, left=anchor.left, width=anchor.width, height=anchor.height)

    def run(self, specific_test=None, wait_callback=None, check_abort=None, item_callback=None):
        # 支援傳入單一測項名稱、測項清單，或預設執行 Enable == YES 者
        if specific_test is not None:
            if isinstance(specific_test, (list, tuple, set)):
                active_items = self.df_item[self.df_item["Test_Item"].isin(specific_test)]
            else:
                active_items = self.df_item[self.df_item["Test_Item"] == specific_test]
        else:
            active_items = self.df_item[self.df_item["Enable"] == "YES"]
            
        if active_items.empty:
            self.log(f"[!] 找不到啟用的測項: {specific_test if specific_test else 'NO ENABLED ITEMS'}")
            return
        
        for _, item in active_items.iterrows():
            if check_abort and check_abort():
                self.log("[*] 收到終止訊號，停止測試排程。")
                break

            test_name = item["Test_Item"]
            target_sheet = item["Report_Sheet"]
            setup_file = item["Setup_File"]
            self.log(f"\n==============================================")
            self.log(f"          開始執行測項: {test_name}")
            self.log(f"==============================================")
            
            if item_callback:
                try: item_callback(test_name, "RUNNING")
                except Exception: pass

            try:
                # Step 1: 參數預先評估
                vars_dict = self.evaluate_variables(test_name)
                self.log(f"[*] 解析完成此測項的所有變數:")
                for k, v in vars_dict.items():
                    self.log(f"    - {k} = {v}")
                    
                # Step 2: 載入示波器基礎設定
                self.control_oscilloscope(setup_file)
                
                # Step 3: 分階段迴圈執行 Layout Mapping 表定義的任務
                layouts = self.df_layout[self.df_layout["Test_Item"] == test_name].sort_values("Phase_Step")
                
                for _, step in layouts.iterrows():
                    if check_abort and check_abort():
                        self.log("[*] 收到終止訊號，停止當前測項。")
                        break

                    prompt_name = step["Prompt_Name"]
                    required_load = vars_dict.get(step["Load_Variable"], "0.0")
                    pic_target = step["Picture_Anchor"]
                    
                    meas_types_str = str(step.get("Meas_Types", "PK2pk"))
                    meas_types = [t for t in meas_types_str.split(",") if t.strip()]
                    
                    data_target_str = str(step["Vpp_Data_Cell"])
                    data_targets = [c for c in data_target_str.split(",") if c.strip()]
                    
                    self.log("\n----------------------------------------------")
                    self.log(f"[等待操作 - Phase {step['Phase_Step']}: {prompt_name}]")
                    self.log(f"請將電子負載設定為: {required_load} (安培)")
                    self.log("波形穩定後，按 [Enter] 或點擊 [確認/下一步] 擷取波形...")
                    self.log("----------------------------------------------")
                    
                    step_info = {
                        "test_item": test_name,
                        "phase_step": step["Phase_Step"],
                        "prompt_name": prompt_name,
                        "required_load": required_load,
                        "pic_target": pic_target,
                        "data_targets": data_targets,
                        "meas_types": meas_types
                    }

                    if wait_callback:
                        # 呼叫 GUI 或 AI 提供的等待確認回呼
                        wait_callback(step_info)
                    else:
                        input()
                    
                    if check_abort and check_abort():
                        self.log("[*] 收到終止訊號，取消本次截圖與回填。")
                        break

                    self.log(f"[*] 截圖並回填 Excel ({pic_target}, {data_targets})...")
                    vpp_list, img = self.capture_waveform(f"{test_name}_Phase{step['Phase_Step']}.png", meas_types)
                    self.write_to_report(target_sheet, vpp_list, img, pic_target, data_targets, meas_types)

                if item_callback:
                    try: item_callback(test_name, "DONE")
                    except Exception: pass

            except Exception as e:
                self.log(f"[!] 執行測項 {test_name} 發生錯誤: {e}")
                if item_callback:
                    try: item_callback(test_name, "ERROR")
                    except Exception: pass

        self.log("\n[*] 測試清單已處理完畢！儲存報告中...")
        try:
            self.wb_report.save()
            self.log("[*] 報告儲存成功！")
        except Exception as e:
            self.log(f"[!] 報告儲存失敗，請確認檔案是否被其他程式鎖定: {e}")

if __name__ == "__main__":
    target = None
    if len(sys.argv) > 1:
        target = sys.argv[1]
        
    engine = AutoTestEngine()
    engine.run(specific_test=target)
