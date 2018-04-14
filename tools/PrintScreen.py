# coding = utf-8
'''
Created on 2017年7月31日

@author: server
'''

from ctypes import Structure, sizeof, memmove, pointer, memset
from ctypes.wintypes import WORD, DWORD, LONG
import win32clipboard
import sys, time
import win32api, win32con

class BITMAPFILEHEADER(Structure):    # https://msdn.microsoft.com/en-us/library/windows/desktop/dd183374(v=vs.85).aspx
    _pack_   = 1                      # structure field byte alignment
    _fields_ = [
        ('bfType',      WORD),
        ('bfSize',      DWORD),
        ('bfReserved1', WORD),
        ('bfReserved2', WORD),
        ('bfOffBits',   DWORD),
        ]   
SIZEOF_BITMAPFILEHEADER = sizeof(BITMAPFILEHEADER)

class BITMAPINFOHEADER(Structure):          # https://msdn.microsoft.com/en-us/library/windows/desktop/dd183376(v=vs.85).aspx
    _pack_   = 1
    _fields_ = [
        ('biSize',          DWORD),
        ('biWidth',         LONG),
        ('biHeight',        LONG),
        ('biPLanes',        WORD),
        ('biBitCount',      WORD),
        ('biCompression',   DWORD),
        ('biSizeImage',     DWORD),
        ('biXpelsPerMeter', LONG),
        ('biYpelsPerMeter', LONG),
        ('biClrUsed',       DWORD),
        ('biClrImportant',  DWORD),
        ]
SIZEOF_BITMAPINFOHEADER = sizeof(BITMAPINFOHEADER)

def printScreen(filename):
    time.sleep(5)
    try:
        win32api.keybd_event(0x91, 0, 0, 0)    # https://msdn.microsoft.com/en-us/library/windows/desktop/dd375731(v=vs.85).aspx   0x91 --> win key
        win32api.keybd_event(0x2C, 0, 0, 0)    # 0x2C --> PRINT SCREEN key
        win32api.keybd_event(0x91, 0, win32con.KEYEVENTF_KEYUP, 0)    # https://msdn.microsoft.com/en-us/library/windows/desktop/ms646304(v=vs.85).aspx
        win32api.keybd_event(0x2C, 0, win32con.KEYEVENTF_KEYUP, 0)
    except:
        print('keyboard event does not successful.')
        sys.exit(1)
    
    BI_BITFIELDS = 3
    
    win32clipboard.OpenClipboard()    # https://msdn.microsoft.com/zh-cn/library/windows/desktop/ff468802(v=vs.85).aspx
    try:
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):    # https://msdn.microsoft.com/en-us/library/windows/desktop/ms649013(v=vs.85).aspx
            data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
        else:
            print('cliboard does not contain an image in DIB format.')
            sys.exit(1)
    finally:
        win32clipboard.CloseClipboard()
        
    BitMapInfoHeaderHandle = BITMAPINFOHEADER()
    memmove(pointer(BitMapInfoHeaderHandle), data, SIZEOF_BITMAPINFOHEADER)
    
    if BitMapInfoHeaderHandle.biCompression != BI_BITFIELDS:
        print('insupported compression type {}'.format(BitMapInfoHeaderHandle.biCompression))
        sys.exit(1)
        
    BitMapFileHeaderHandle = BITMAPFILEHEADER()
    memset(pointer(BitMapFileHeaderHandle), 0, SIZEOF_BITMAPFILEHEADER)
    BitMapFileHeaderHandle.bfType = ord('B') | (ord('M') << 8)
    BitMapFileHeaderHandle.bfSize = SIZEOF_BITMAPFILEHEADER + len(data)
    SIZEOF_COLORTABLE = 0
    BitMapFileHeaderHandle.bfOffBits = SIZEOF_BITMAPFILEHEADER + SIZEOF_BITMAPINFOHEADER + SIZEOF_COLORTABLE
    
    with open(filename, 'wb') as bmp_file:
        bmp_file.write(BitMapFileHeaderHandle)
        bmp_file.write(data)
    print('file "{}" created from clipboard image'.format(filename))
    
if __name__ == '__main__':
    printScreen('test.png')
