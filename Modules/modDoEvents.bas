Attribute VB_Name = "DoEvents"
Option Explicit

Public Declare Function DoEvents_PeekMessage _
               Lib "Coredll" _
               Alias "PeekMessageW" (ByVal lpMsg As String, _
                                     ByVal hWnd As Long, _
                                     ByVal wMsgFilterMin As Long, _
                                     ByVal wMsgFilterMax As Long, _
                                     ByVal wRemoveMsg As Long) As Long

Public Declare Function DoEvents_GetMessage _
               Lib "Coredll" _
               Alias "GetMessageW" (ByVal lpMsg As String, _
                                    ByVal hWnd As Long, _
                                    ByVal wMsgFilterMin As Long, _
                                    ByVal wMsgFilterMax As Long) As Long

Public Declare Function DoEvents_TranslateMessage _
               Lib "Coredll" _
               Alias "TranslateMessage" (ByVal lpMsg As String) As Long

Public Declare Function DoEvents_DispatchMessage _
               Lib "Coredll" _
               Alias "DispatchMessageW" (ByVal lpMsg As String) As Long

Public Sub DoEvents_Run()

    Dim strMsg As String 'MSG

    strMsg = DoEvents_MakeMSG(0, 0, 0, 0, 0, 0)
   
    Do While DoEvents_PeekMessage(strMsg, 0, 0, 0, True)
        DoEvents_TranslateMessage strMsg
        DoEvents_DispatchMessage strMsg
    Loop

End Sub

Private Function DoEvents_MakeMSG(ByVal hWnd As Long, _
                                  ByVal message As Long, _
                                  ByVal wParam As Long, _
                                  ByVal lParam As Long, _
                                  ByVal time As Long, _
                                  ByVal pt As Long) As String

    Dim varMembers As Variant

    varMembers = Array(UDTHelper_ToBinaryString(CLng(hWnd), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(message), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(wParam), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(lParam), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(time), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(pt), UDTHelper_CE_LONG))

    DoEvents_MakeMSG = Join(varMembers, vbNullString)

End Function



