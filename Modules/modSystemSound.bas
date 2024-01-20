Attribute VB_Name = "SystemSound"
Option Explicit

'Private Declares are not supported for Windows CE.

Public Declare Function SystemSound_PlaySound _
               Lib "Coredll" _
               Alias "PlaySoundW" (ByVal lpszName As String, _
                                   ByVal hModule As Long, _
                                   ByVal dwFlags As Long) As Long

Private Const SystemSound_SND_ASYNC = &H1         'Play asynchronously

Private Const SystemSound_SND_ALIAS = &H10000     'Name is a WIN.INI [sounds] entry

Public Const SystemSound_ceSystemSoundAsterisk       As String = "SystemAsterisk"

Public Const SystemSound_ceSystemSoundDefault        As String = "SystemDefault"

Public Const SystemSound_ceSystemSoundExclamation    As String = "SystemExclamation"

Public Const SystemSound_ceSystemSoundSystemExit     As String = "SystemExit"

Public Const SystemSound_ceSystemSoundSystemHand     As String = "SystemHand"

Public Const SystemSound_ceSystemSoundSystemQuestion As String = "SystemQuestion"

Public Const SystemSound_ceSystemSoundSystemStart    As String = "SystemStart"

Public Const SystemSound_ceSystemSoundSystemWelcome  As String = "SystemWelcome"

Public Function SystemSound_Play(ByVal Id As String) As Long
    SystemSound_Play = SystemSound_PlaySound(Id, 0, SystemSound_SND_ALIAS Or SystemSound_SND_ASYNC)
End Function



