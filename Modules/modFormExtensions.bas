Attribute VB_Name = "FormExtensions"
Option Explicit
                                       
Public Declare Function FormExtensions_SetWindowPos _
               Lib "Coredll" _
               Alias "SetWindowPos" (ByVal hWnd As Long, _
                                     ByVal hWndInsertAfter As Long, _
                                     ByVal x As Long, _
                                     ByVal y As Long, _
                                     ByVal cx As Long, _
                                     ByVal cy As Long, _
                                     ByVal wFlags As Long) As Long

Public Const FormExtensions_FLAGS = 3 'SWP_NOMOVE Or SWP_NOSIZE

Public Const FormExtensions_HWND_TOPMOST = -1

Public Const FormExtensions_HWND_NOTOPMOST = -2

Public Function FormExtensions_SetTopMostState(Form As Form, State As Boolean) As Long

    If State Then

        FormExtensions_SetTopMostState = FormExtensions_SetWindowPos(Form.hWnd, FormExtensions_HWND_TOPMOST, 0, 0, 0, 0, FormExtensions_FLAGS)
    Else

        FormExtensions_SetTopMostState = FormExtensions_SetWindowPos(Form.hWnd, FormExtensions_HWND_NOTOPMOST, 0, 0, 0, 0, FormExtensions_FLAGS)
    End If
     
End Function



