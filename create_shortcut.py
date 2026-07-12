import os
import win32com.client
from winreg import OpenKey, QueryValueEx, HKEY_CURRENT_USER

def get_desktop_path():
    try:
        key = OpenKey(HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
        return QueryValueEx(key, "Desktop")[0]
    except Exception:
        return os.path.join(os.environ["USERPROFILE"], "Desktop")

desktop = get_desktop_path()
path = os.path.join(desktop, "Recoil Control.lnk")
target = r"m:\Projetos DevDuo\Spray Control\dist\Recoil Control.exe"
wDir = r"m:\Projetos DevDuo\Spray Control"

shell = win32com.client.Dispatch('WScript.Shell')
shortcut = shell.CreateShortCut(path)
shortcut.Targetpath = target
shortcut.WorkingDirectory = wDir
shortcut.save()
print("Shortcut created at:", path)
