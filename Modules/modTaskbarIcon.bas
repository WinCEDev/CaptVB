Attribute VB_Name = "TaskbarIcon"
Option Explicit

Public Declare Function TaskbarIcon_ImageList_GetIcon _
               Lib "Coredll" _
               Alias "ImageList_GetIcon" (ByVal himl As Long, _
                                          ByVal i As Long, _
                                          ByVal Flags As Long) As Long

Public Declare Function TaskbarIcon_DestroyIcon _
               Lib "Coredll" _
               Alias "DestroyIcon" (ByVal hIcon As Long) As Long

Public Declare Function TaskbarIcon_SendMessage _
               Lib "Coredll" _
               Alias "SendMessageW" (ByVal hWnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     ByVal lParam As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Icon size constants.

Public Const TaskbarIcon_tbSmallIcon As Long = 0 'ICON_SMALL

Public Const TaskbarIcon_tbLargeIcon As Long = 1 'ICON_BIG

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Window Messages.

Private Const TaskbarIcon_WM_SETICON As Long = &H80

Public Function TaskbarIcon_Let(ByVal FormHandle As Long, _
                                ByVal ImageListHandle As Long, _
                                ByVal Index As Long, _
                                ByVal IconSize As Long) As Long

    Dim lngIcon As Long

    lngIcon = TaskbarIcon_ImageList_GetIcon(ImageListHandle, Index, 0)
    TaskbarIcon_SendMessage FormHandle, TaskbarIcon_WM_SETICON, IconSize, lngIcon

    TaskbarIcon_Let = lngIcon

End Function

Public Function TaskbarIcon_Destroy(ByVal Icon As Long) As Long
    TaskbarIcon_Destroy = TaskbarIcon_DestroyIcon(Icon)
End Function



