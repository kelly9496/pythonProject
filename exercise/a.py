import win32gui,win32ui,win32con
print('sucessful')
def get_windows(windowsname,filename):
    handle = win32gui.FindWindow(None,windowsname)
    win32gui.SetForegroundWindow(handle)
    hdDC = win32gui.GetWindowDC(handle)
    newhdDC = win32ui.CreateDCFromHandle(hdDC)
    saveDC = newhdDC.CreateCompatibleDC()
    saveBitmap = win32ui.CreateBitmap()
    left, top, right, bottom = win32gui.GetWindowRect(handle)
    width = right - left
    height = bottom - top
    saveBitmap.CreateCompatibleBitmap(newhdDC, width, height)
    saveDC.SelectObject(saveBitmap)
    saveDC.BitBlt((0, 0), (width, height), newhdDC, (0, 0), win32con.SRCCOPY)
    saveBitmap.SaveBitmapFile(saveDC, filename)
get_windows("PyWin32","截图.png")