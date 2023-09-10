Attribute VB_Name = "NotifyIcon"
Option Explicit

Public Declare Function NotifyIcon_Shell_NotifyIcon _
               Lib "Coredll" _
               Alias "Shell_NotifyIcon" (ByVal dwMessage As Long, _
                                         ByVal pnid As String) As Long

Public Declare Function NotifyIcon_ImageList_GetIcon _
               Lib "Coredll" _
               Alias "ImageList_GetIcon" (ByVal himl As Long, _
                                          ByVal i As Long, _
                                          ByVal Flags As Long) As Long

Public Declare Function NotifyIcon_DestroyIcon _
               Lib "Coredll" _
               Alias "DestroyIcon" (ByVal hIcon As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NotifyIcon Messages.
'https://learn.microsoft.com/en-us/previous-versions/ms942613(v=msdn.10)

Private Const NotifyIcon_NIM_ADD         As Long = &H0 'Adds an icon to the status area.

Private Const NotifyIcon_NIM_MODIFY      As Long = &H1 'Modifies an icon in the status area.

Private Const NotifyIcon_NIM_DELETE      As Long = &H2 'Deletes an icon from the status area.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NotifyIcon Flags.
'https://learn.microsoft.com/en-us/previous-versions/ms961260(v=msdn.10)

Private Const NotifyIcon_NIF_MESSAGE     As Long = &H1 'The uCallbackMessage member is valid.

Private Const NotifyIcon_NIF_ICON        As Long = &H2 'The hIcon member is valid.

Private Const NotifyIcon_NIF_TIP         As Long = &H4 'The szTip member is valid.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Window Messages.

Public Const NotifyIcon_WM_LBUTTONDBLCLK As Long = &H203 'Double-click.

Public Const NotifyIcon_WM_LBUTTONDOWN   As Long = &H201 'Button down.

Public Const NotifyIcon_WM_LBUTTONUP     As Long = &H202 'Button up.

Public Const NotifyIcon_WM_RBUTTONUP     As Long = &H205 'Button up.

Public Const NotifyIcon_WM_KEYDOWN = &H100

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Used by 'NotifyIcon_Add' and 'NotifyIcon_Remove'.

Private Const NotifyIcon_NOTIFYICONDATA_LEN As Long = 24 'Length in bytes of the NOTIFYICONDATA structure.

Private Const NotifyIcon_ICON_ID            As Long = 13 'Unique ID for this icon, since this module only lets you add a single icon, this value is hardcoded. Values 0-12 are reserved: https://learn.microsoft.com/en-us/previous-versions/windows/embedded/ms911889(v=msdn.10).

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Module-level variables.

Private NotifyIcon_CallbackFormHandle       As Long 'Holds the handle of the form receiving callbacks.

Private NotifyIcon_FormImageListHandle      As Long 'Holds the handle of the form's image list.

Private NotifyIcon_IconHandle               As Long 'Holds the handle of the icon used for the system tray.

Public Function NotifyIcon_Add(ByVal FormHandle As Long, _
                               ByVal ImageListHandle As Long, _
                               ByVal Index As Long) As Long
                               
    If NotifyIcon_CallbackFormHandle <> 0 Then 'An icon is already present.
        NotifyIcon_Remove 'First remove the old icon.
    End If

    NotifyIcon_CallbackFormHandle = FormHandle
    NotifyIcon_FormImageListHandle = ImageListHandle
    NotifyIcon_IconHandle = NotifyIcon_ImageList_GetIcon(NotifyIcon_FormImageListHandle, Index, 0)
    
    Dim strNOTIFYICONDATA As String
    
    strNOTIFYICONDATA = NotifyIcon_MakeNOTIFYICONDATA(NotifyIcon_NOTIFYICONDATA_LEN, NotifyIcon_CallbackFormHandle, NotifyIcon_ICON_ID, NotifyIcon_NIF_ICON Or NotifyIcon_NIF_MESSAGE, NotifyIcon_WM_LBUTTONDOWN, NotifyIcon_IconHandle)

    NotifyIcon_Add = NotifyIcon_Shell_NotifyIcon(NotifyIcon_NIM_ADD, strNOTIFYICONDATA)

End Function

Public Function NotifyIcon_Remove() As Long

    If NotifyIcon_CallbackFormHandle <> 0 Then

        Dim strNOTIFYICONDATA As String

        strNOTIFYICONDATA = NotifyIcon_MakeNOTIFYICONDATA(NotifyIcon_NOTIFYICONDATA_LEN, NotifyIcon_CallbackFormHandle, NotifyIcon_ICON_ID, 0, 0, 0)
    
        NotifyIcon_Remove = NotifyIcon_Shell_NotifyIcon(NotifyIcon_NIM_DELETE, strNOTIFYICONDATA) And NotifyIcon_DestroyIcon(NotifyIcon_IconHandle)

        NotifyIcon_CallbackFormHandle = 0
        NotifyIcon_FormImageListHandle = 0
        NotifyIcon_IconHandle = 0

    End If

End Function

Public Function NotifyIcon_Modify(ByVal Index As Long) As Long

    If NotifyIcon_CallbackFormHandle <> 0 Then

        Dim lngNewIcon As Long

        lngNewIcon = NotifyIcon_ImageList_GetIcon(NotifyIcon_FormImageListHandle, Index, 0) 'Load new icon.

        Dim strNOTIFYICONDATA As String

        strNOTIFYICONDATA = NotifyIcon_MakeNOTIFYICONDATA(NotifyIcon_NOTIFYICONDATA_LEN, NotifyIcon_CallbackFormHandle, NotifyIcon_ICON_ID, NotifyIcon_NIF_ICON Or NotifyIcon_NIF_MESSAGE, NotifyIcon_WM_LBUTTONDOWN, lngNewIcon)

        NotifyIcon_Modify = NotifyIcon_Shell_NotifyIcon(NotifyIcon_NIM_MODIFY, strNOTIFYICONDATA)
        
        'Delete old icon and load the new icon into the NotifyIcon_IconHandle variable.
        NotifyIcon_DestroyIcon NotifyIcon_IconHandle
        NotifyIcon_IconHandle = lngNewIcon

    End If

End Function

Private Function NotifyIcon_MakeNOTIFYICONDATA(ByVal cbSize As Long, _
                                               ByVal hWnd As Long, _
                                               ByVal uID As Long, _
                                               ByVal uFlags As Long, _
                                               ByVal uCallbackMessage As Long, _
                                               ByVal hIcon As Long) As String

    Dim varMembers As Variant

    varMembers = Array(UDTHelper_ToBinaryString(CLng(cbSize), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(hWnd), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(uID), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(uFlags), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(uCallbackMessage), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(hIcon), UDTHelper_CE_LONG))

    NotifyIcon_MakeNOTIFYICONDATA = Join(varMembers, vbNullString)

End Function



