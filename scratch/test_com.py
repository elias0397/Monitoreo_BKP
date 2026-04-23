import win32com.client
import sys

def test_com(progid):
    try:
        obj = win32com.client.Dispatch(progid)
        print(f"[OK] {progid} connected")
        return True
    except Exception as e:
        print(f"[ERR] {progid} failed: {e}")
        return False

if __name__ == "__main__":
    print(f"Python version: {sys.version}")
    test_com("Outlook.Application")
    test_com("Excel.Application")
    test_com("Word.Application")
    test_com("Shell.Application")
