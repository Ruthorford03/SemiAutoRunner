"""
SCPI Controller (標準 VISA 儀器通訊控制器)
==============================================
提供基於 pyvisa 的標準 LAN/TCPIP SCPI 連線、指令寫入與查詢功能。
不包含任何示波器螢幕操作或逆向控制代碼。
"""
import threading

try:
    import pyvisa
    HAS_VISA = True
except ImportError:
    HAS_VISA = False

class SCPIController:
    """ 透過 pyvisa 傳送標準 SCPI 指令控制示波器 """

    def __init__(self, ip):
        self.ip = ip
        self.scope = None
        self.rm = None
        self.connected = False
        self.lock = threading.Lock()

    def connect(self):
        if not HAS_VISA:
            return False, "pyvisa not installed"
        try:
            self.rm = pyvisa.ResourceManager()
            addr = f"TCPIP::{self.ip}::INSTR"
            self.scope = self.rm.open_resource(addr, timeout=30000)
            idn = self.scope.query("*IDN?").strip()
            self.connected = True
            return True, idn
        except Exception as e:
            self.connected = False
            return False, str(e)

    def write(self, cmd):
        if not self.connected:
            return False
        with self.lock:
            try:
                self.scope.write(cmd)
                return True
            except Exception:
                return False

    def query(self, cmd):
        if not self.connected:
            return None
        with self.lock:
            try:
                return self.scope.query(cmd).strip()
            except Exception:
                return None

    def close(self):
        try:
            if self.scope:
                self.scope.close()
            if self.rm:
                self.rm.close()
        except Exception:
            pass
        self.connected = False
