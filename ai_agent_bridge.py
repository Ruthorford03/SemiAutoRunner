"""
AI Agent Bridge (AI 代理人操作介面與橋接模組)
==============================================
本模組為 SemiAutoRunner 系統預留給未來 AI (如 Gemini Vision / 視覺代理人)
進行自主測試、即時波形判讀、SCPI 遙控與流程調度的專屬 API 接口。
"""
import io
import base64
from PIL import Image

class AIAgentBridge:
    def __init__(self, gui_app=None, engine=None):
        self.gui = gui_app
        self.engine = engine
        self.auto_pilot_enabled = False
        self.ai_hooks = []

    def bind_contexts(self, gui_app=None, engine=None):
        """ 綁定 GUI 實例與測試引擎實例 """
        if gui_app: self.gui = gui_app
        if engine: self.engine = engine

    # ================= 視覺與畫面擷取 API =================
    def get_screen_image(self) -> Image.Image:
        """
        獲取示波器當前最新幀畫面 (PIL Image 格式)
        供未來 AI 視覺模型 (如 Gemini Vision) 分析波形品質、觸發是否成功
        """
        if self.gui and hasattr(self.gui, "current_pil_image") and self.gui.current_pil_image:
            return self.gui.current_pil_image.copy()
        
        # 若 GUI 未提供，嘗試直接從示波器 HTTP 獲取
        if self.gui and hasattr(self.gui, "scope_ip"):
            import requests
            try:
                r = requests.get(f"http://{self.gui.scope_ip}/image.png", timeout=3)
                r.raise_for_status()
                return Image.open(io.BytesIO(r.content))
            except Exception:
                pass
        return None

    def get_screen_base64(self, format="PNG") -> str:
        """ 獲取當前示波器畫面的 Base64 字串，方便直接嵌入 Prompt 送給多模態 LLM """
        img = self.get_screen_image()
        if img is None:
            return ""
        buf = io.BytesIO()
        img.save(buf, format=format)
        return base64.b64encode(buf.getvalue()).decode("utf-8")

    # ================= 示波器 SCPI 遙控 API =================
    def send_scpi_command(self, cmd: str) -> bool:
        """ 供 AI 執行示波器控制指令 (例如調整時基、垂直刻度、觸發電位) """
        if self.engine and hasattr(self.engine, "scpi") and self.engine.scpi:
            return self.engine.scpi.write(cmd)
        elif self.gui and hasattr(self.gui, "scpi") and self.gui.scpi:
            return self.gui.scpi.write(cmd)
        return False

    def query_scpi(self, cmd: str) -> str:
        """ 供 AI 查詢示波器狀態或量測值 """
        if self.engine and hasattr(self.engine, "scpi") and self.engine.scpi:
            return self.engine.scpi.query(cmd)
        elif self.gui and hasattr(self.gui, "scpi") and self.gui.scpi:
            return self.gui.scpi.query(cmd)
        return None

    # ================= 流程調度與確認 API =================
    def get_current_step_info(self) -> dict:
        """ 獲取當前測試階段的提示、所需負載值、錨點等資訊 """
        if self.gui and hasattr(self.gui, "current_step_info"):
            return self.gui.current_step_info
        return None

    def confirm_step(self):
        """ 供 AI 程式在確認波形穩定或儀器設定完成後，自動按下「確認/下一步」 """
        if self.gui and hasattr(self.gui, "on_confirm_next"):
            self.log("[AI Agent] 觸發自動確認步驟 (confirm_step)")
            self.gui.on_confirm_next()

    def abort_test(self):
        """ 供 AI 在偵測到異常、超標或緊急狀況時中止測試流程 """
        if self.gui and hasattr(self.gui, "on_stop_test"):
            self.log("[AI Agent] 發出測試中止命令 (abort_test)")
            self.gui.on_stop_test()

    def start_tests(self, items: list = None):
        """ 供 AI 啟動指定測項清單 """
        if self.gui and hasattr(self.gui, "start_execution"):
            if items:
                self.gui.set_selected_items(items)
            self.gui.start_execution()

    def get_test_status(self) -> dict:
        """ 獲取所有測項當前的測試狀態 (PENDING, RUNNING, DONE, ERROR) """
        if self.gui and hasattr(self.gui, "item_status"):
            return dict(self.gui.item_status)
        return {}

    def log(self, message: str):
        """ 在 GUI 與終端機輸出 AI Agent 決策日誌 """
        msg_formatted = f"[AI] {message}"
        if self.gui and hasattr(self.gui, "append_log"):
            self.gui.append_log(msg_formatted)
        print(msg_formatted)
