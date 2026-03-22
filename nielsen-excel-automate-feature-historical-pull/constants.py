import enum
import os


APPIUM_HOST = "http://127.0.0.1"
APPIUM_PORT = 10000
WAD_PORT = 4724
WAD_EXE = r"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe"
ROOT_DIR = os.path.dirname(__file__)


class AppiumServerState(enum.Enum):
    STARTING = "starting"
    LISTENING = "listening"
    TIMED_OUT = "timed_out"
