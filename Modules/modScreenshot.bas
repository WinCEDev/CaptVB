Attribute VB_Name = "Screenshot"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GDI Declarations.

Public Declare Function Screenshot_CreateDC _
               Lib "Coredll" _
               Alias "CreateDCW" (ByVal lpDriverName As String, _
                                  ByVal lpDeviceName As String, _
                                  ByVal lpOutput As String, _
                                  ByRef lpInitData As Long) As Long
                                  
Public Declare Function Screenshot_DeleteDC _
               Lib "Coredll" _
               Alias "DeleteDC" (ByVal hdc As Long) As Long

Public Declare Function Screenshot_CreateCompatibleDC _
               Lib "Coredll" _
               Alias "CreateCompatibleDC" (ByVal hdc As Long) As Long

Public Declare Function Screenshot_SelectObject _
               Lib "Coredll" _
               Alias "SelectObject" (ByVal hdc As Long, _
                                     ByVal hObject As Long) As Long

Public Declare Function Screenshot_BitBlt _
               Lib "Coredll" _
               Alias "BitBlt" (ByVal hDestDC As Long, _
                               ByVal x As Long, _
                               ByVal y As Long, _
                               ByVal nWidth As Long, _
                               ByVal nHeight As Long, _
                               ByVal hSrcDC As Long, _
                               ByVal xSrc As Long, _
                               ByVal ySrc As Long, _
                               ByVal dwRop As Long) As Long

Public Declare Function Screenshot_DeleteObject _
               Lib "Coredll" _
               Alias "DeleteObject" (ByVal hObject As Long) As Long

Public Declare Function Screenshot_CreateDIBSection _
               Lib "Coredll" _
               Alias "CreateDIBSection" (ByVal hdc As Long, _
                                         ByVal pbmi As String, _
                                         ByVal iUsage As Long, _
                                         ByRef ppvBits As Long, _
                                         ByVal hSection As Long, _
                                         ByVal dwOffset As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'File declarations.

Public Declare Function Screenshot_CreateFile _
               Lib "Coredll" _
               Alias "CreateFileW" (ByVal lpFileName As String, _
                                    ByVal dwDesiredAccess As Long, _
                                    ByVal dwShareMode As Long, _
                                    lpSecurityAttributes As Long, _
                                    ByVal dwCreationDisposition As Long, _
                                    ByVal dwFlagsAndAttributes As Long, _
                                    ByVal hTemplateFile As Long) As Long

Public Declare Function Screenshot_WriteFileLong _
               Lib "Coredll" _
               Alias "WriteFile" (ByVal hFile As Long, _
                                  ByRef lpBuffer As Long, _
                                  ByVal nNumberOfBytesToWrite As Long, _
                                  ByRef lpNumberOfBytesWritten As Long, _
                                  ByVal lpOverlapped As Long) As Long

Public Declare Function Screenshot_WriteFileInteger _
               Lib "Coredll" _
               Alias "WriteFile" (ByVal hFile As Long, _
                                  ByRef lpBuffer As Integer, _
                                  ByVal nNumberOfBytesToWrite As Long, _
                                  ByRef lpNumberOfBytesWritten As Long, _
                                  ByVal lpOverlapped As Long) As Long

Public Declare Function Screenshot_WriteFileString _
               Lib "Coredll" _
               Alias "WriteFile" (ByVal hFile As Long, _
                                  ByRef lpBuffer As String, _
                                  ByVal nNumberOfBytesToWrite As Long, _
                                  ByRef lpNumberOfBytesWritten As Long, _
                                  ByVal lpOverlapped As Long) As Long

Public Declare Function Screenshot_WriteFilePointer _
               Lib "Coredll" _
               Alias "WriteFile" (ByVal hFile As Long, _
                                  ByVal lpBuffer As Long, _
                                  ByVal nNumberOfBytesToWrite As Long, _
                                  ByRef lpNumberOfBytesWritten As Long, _
                                  ByVal lpOverlapped As Long) As Long

Public Declare Function Screenshot_CloseHandle _
               Lib "Coredll" _
               Alias "CloseHandle" (ByVal hObject As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BitBlt constants.

Private Const Screenshot_SRCCOPY               As Long = &HCC0020

Private Const Screenshot_DIB_RGB_COLORS        As Long = &H0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'File constants.

Private Const Screenshot_GENERIC_WRITE         As Long = &H40000000

Private Const Screenshot_FILE_ATTRIBUTE_NORMAL As Long = &H80

Private Const Screenshot_CREATE_ALWAYS         As Long = 2

Public Function Screenshot_Take(ByVal FilePath As String) As Boolean

    Dim lngWidth As Long, lngHeight As Long

    lngWidth = CLng(Screen.Width \ Screen.TwipsPerPixelX)
    lngHeight = CLng(Screen.Height \ Screen.TwipsPerPixelY)
    
    Dim lngScreenDC             As Long 'HDC

    Dim lngMemoryDC             As Long 'HDC

    Dim lngBitmapHandle         As Long 'HBITMAP

    Dim lngOriginalBitmapHandle As Long 'HBITMAP

    lngScreenDC = Screenshot_CreateDC("DISPLAY", vbNullString, vbNullString, vbnullptr)
    lngMemoryDC = Screenshot_CreateCompatibleDC(lngScreenDC)
    
    Dim strBITMAPINFO As String 'BITMAPINFO Structure.
    
    strBITMAPINFO = Screenshot_MakeBITMAPINFO(Screenshot_Screenshot_MakeBITMAPINFOHEADER(40, lngWidth, lngHeight, 1, 24, 0, 0, 3780, 3780, 0, 0), 0)
    
    Dim lngBits As Long 'Pointer to image data.

    lngBitmapHandle = Screenshot_CreateDIBSection(lngMemoryDC, strBITMAPINFO, Screenshot_DIB_RGB_COLORS, lngBits, 0, 0)
    lngOriginalBitmapHandle = Screenshot_SelectObject(lngMemoryDC, lngBitmapHandle)
    
    Screenshot_BitBlt lngMemoryDC, 0, 0, lngWidth, lngHeight, lngScreenDC, 0, 0, Screenshot_SRCCOPY

    lngBitmapHandle = Screenshot_SelectObject(lngMemoryDC, lngOriginalBitmapHandle)
    
    Screenshot_DeleteDC lngScreenDC
    Screenshot_DeleteDC lngMemoryDC
    
    Dim lngFile As Long 'FILE

    lngFile = Screenshot_CreateFile(FilePath, Screenshot_GENERIC_WRITE, 0, vbnullptr, Screenshot_CREATE_ALWAYS, Screenshot_FILE_ATTRIBUTE_NORMAL, 0)
    
    Dim lngImageSize As Long

    lngImageSize = lngWidth * lngHeight * 3
    
    Dim lngNumberOfBytesWritten As Long

    lngNumberOfBytesWritten = CLng(0)
    
    'Write BITMAPFILEHEADER.
    Screenshot_WriteFileInteger lngFile, &H4D42, UDTHelper_CE_INTEGER, lngNumberOfBytesWritten, 0 'bfType
    Screenshot_WriteFileLong lngFile, lngImageSize + 54, UDTHelper_CE_LONG, lngNumberOfBytesWritten, 0 'bfSize
    Screenshot_WriteFileInteger lngFile, 0, UDTHelper_CE_INTEGER, lngNumberOfBytesWritten, 0 'bfReserved1
    Screenshot_WriteFileInteger lngFile, 0, UDTHelper_CE_INTEGER, lngNumberOfBytesWritten, 0 'bfReserved2
    Screenshot_WriteFileLong lngFile, 54, UDTHelper_CE_LONG, lngNumberOfBytesWritten, 0 'bfOffBits
    
    'Write BITMAPINFOHEADER.
    Screenshot_WriteFileLong lngFile, 40, UDTHelper_CE_LONG, lngNumberOfBytesWritten, 0 'biSize
    Screenshot_WriteFileLong lngFile, lngWidth, UDTHelper_CE_LONG, lngNumberOfBytesWritten, 0 'biWidth
    Screenshot_WriteFileLong lngFile, lngHeight, UDTHelper_CE_LONG, lngNumberOfBytesWritten, 0 'biHeight
    Screenshot_WriteFileInteger lngFile, 1, UDTHelper_CE_INTEGER, lngNumberOfBytesWritten, 0 'biPlanes
    Screenshot_WriteFileInteger lngFile, 24, UDTHelper_CE_INTEGER, lngNumberOfBytesWritten, 0 'biBitCount
    Screenshot_WriteFileLong lngFile, 0, UDTHelper_CE_LONG, lngNumberOfBytesWritten, 0 'biCompression
    Screenshot_WriteFileLong lngFile, 0, UDTHelper_CE_LONG, lngNumberOfBytesWritten, 0 'biSizeImage
    Screenshot_WriteFileLong lngFile, 3780, UDTHelper_CE_LONG, lngNumberOfBytesWritten, 0 'biXPelsPerMeter
    Screenshot_WriteFileLong lngFile, 3780, UDTHelper_CE_LONG, lngNumberOfBytesWritten, 0 'biYPelsPerMeter
    Screenshot_WriteFileLong lngFile, 0, UDTHelper_CE_LONG, lngNumberOfBytesWritten, 0 'biClrUsed
    Screenshot_WriteFileLong lngFile, 0, UDTHelper_CE_LONG, lngNumberOfBytesWritten, 0 'biClrImportant
    
    'Write image data.
    Screenshot_WriteFilePointer lngFile, lngBits, lngImageSize, lngNumberOfBytesWritten, 0 'Image data.
    
    'Close the file.
    Screenshot_CloseHandle lngFile
    
    'Delete the bitmap object.
    Screenshot_DeleteObject lngBitmapHandle

End Function

Private Function Screenshot_MakeBITMAPINFO(bmiHeader As String, _
                                           bmiColors As Long) As String

    Dim varMembers As Variant

    varMembers = Array(bmiHeader, UDTHelper_ToBinaryString(CLng(bmiColors), UDTHelper_CE_LONG))

    Screenshot_MakeBITMAPINFO = Join(varMembers, vbNullString)

End Function

Private Function Screenshot_Screenshot_MakeBITMAPINFOHEADER(biSize As Long, _
                                                            biWidth As Long, _
                                                            biHeight As Long, _
                                                            biPlanes As Integer, _
                                                            biBitCount As Integer, _
                                                            biCompression As Long, _
                                                            biSizeImage As Long, _
                                                            biXPelsPerMeter As Long, _
                                                            biYPelsPerMeter As Long, _
                                                            biClrUsed As Long, _
                                                            biClrImportant As Long) As Long

    Dim varMembers As Variant

    varMembers = Array(UDTHelper_ToBinaryString(CLng(biSize), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(biWidth), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(biHeight), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CInt(biPlanes), UDTHelper_CE_INTEGER), UDTHelper_ToBinaryString(CInt(biBitCount), UDTHelper_CE_INTEGER), UDTHelper_ToBinaryString(CLng(biCompression), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(biSizeImage), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(biXPelsPerMeter), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(biYPelsPerMeter), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(biClrUsed), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(biClrImportant), UDTHelper_CE_LONG))

    Screenshot_Screenshot_MakeBITMAPINFOHEADER = Join(varMembers, vbNullString)

End Function



