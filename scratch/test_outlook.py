import win32com.client
try:
    outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    print("[OK] Outlook connected via EnsureDispatch")
except Exception as e:
    print(f"[ERR] Outlook failed via EnsureDispatch: {e}")
