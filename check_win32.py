
try:
    import win32com.client
    print("win32com is available.")
except ImportError:
    print("win32com is NOT available.")
