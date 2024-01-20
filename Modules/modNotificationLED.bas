Attribute VB_Name = "NotificationLED"
Option Explicit

Public Declare Function NotificationLED_NLedSetDevice _
               Lib "Coredll" _
               Alias "NLedSetDevice" (ByVal nDeviceId As Long, _
                                      ByVal pInput As String) As Long

Public Declare Function NotificationLED_NLedGetDeviceInfo _
               Lib "Coredll" _
               Alias "NLedGetDeviceInfo" (ByVal nInfoId As Long, _
                                          ByVal pOutput As String) As Long

Public Declare Function NotificationLED_NLedGetDeviceInfo_Count _
               Lib "Coredll" _
               Alias "NLedGetDeviceInfo" (ByVal nInfoId As Long, _
                                          ByRef pOutput As Long) As Long

Private Const NotificationLED_NLED_COUNT_INFO_ID    As Long = 0

Private Const NotificationLED_NLED_SUPPORTS_INFO_ID As Long = 1

Private Const NotificationLED_NLED_SETTINGS_INFO_ID As Long = 2

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Flags to be used with NotificationLED_GetSupportedFeatures

Public Const NotificationLED_AdjustTotalCycleTime   As Long = 1

Public Const NotificationLED_AdjustOnTime           As Long = 2

Public Const NotificationLED_AdjustOffTime          As Long = 4

Public Const NotificationLED_MetaCycleOn            As Long = 8

Public Const NotificationLED_MetaCycleOff           As Long = 16

Public Function NotificationLED_Count() As Long

    On Error Resume Next

    NotificationLED_NLedGetDeviceInfo_Count NotificationLED_NLED_COUNT_INFO_ID, NotificationLED_Count
    
    If Err.Number <> 0 Then 'Return 0 if error on API call.
        NotificationLED_Count = 0
    End If

End Function

Public Function NotificationLED_Let(ByVal LedNum As Long, _
                                    ByVal OffOnBlink As Long, _
                                    ByVal TotalCycleTime As Long, _
                                    ByVal OnTime As Long, _
                                    ByVal OffTime As Long, _
                                    ByVal MetaCycleOn As Long, _
                                    ByVal MetaCycleOff As Long) As Long
                                    
    On Error Resume Next
    
    NotificationLED_Let = NotificationLED_NLedSetDevice(NotificationLED_NLED_SETTINGS_INFO_ID, NotificationLED_MakeNLED_SETTINGS_INFO(LedNum, OffOnBlink, TotalCycleTime, OnTime, OffTime, MetaCycleOn, MetaCycleOff))

    If Err.Number <> 0 Then 'Return 0 if error on API call.
        NotificationLED_Let = 0
    End If

End Function

Public Function NotificationLED_GetSupportedFeatures(ByVal LedNum As Long, _
                                                     ByRef CycleAdjust As Long) As Long

    On Error Resume Next

    Dim strLedInfo As String

    strLedInfo = NotificationLED_MakeNLED_SUPPORTS_INFO(LedNum, 0, 0, 0, 0, 0, 0)

    If NotificationLED_NLedGetDeviceInfo(NotificationLED_NLED_SUPPORTS_INFO_ID, strLedInfo) = 1 Then

        CycleAdjust = CLng(UDTHelper_FromBinaryString(MidB(strLedInfo, 5, UDTHelper_CE_LONG)))

        Dim lngResult As Long

        lngResult = CLng(0)

        If CLng(UDTHelper_FromBinaryString(MidB(strLedInfo, 9, UDTHelper_CE_LONG))) = 1 Then
            lngResult = lngResult Or NotificationLED_AdjustTotalCycleTime
        End If

        If CLng(UDTHelper_FromBinaryString(MidB(strLedInfo, 13, UDTHelper_CE_LONG))) = 1 Then
            lngResult = lngResult Or NotificationLED_AdjustOnTime
        End If

        If CLng(UDTHelper_FromBinaryString(MidB(strLedInfo, 17, UDTHelper_CE_LONG))) = 1 Then
            lngResult = lngResult Or NotificationLED_AdjustOffTime
        End If

        If CLng(UDTHelper_FromBinaryString(MidB(strLedInfo, 21, UDTHelper_CE_LONG))) = 1 Then
            lngResult = lngResult Or NotificationLED_MetaCycleOn
        End If

        If CLng(UDTHelper_FromBinaryString(MidB(strLedInfo, 25, UDTHelper_CE_LONG))) = 1 Then
            lngResult = lngResult Or NotificationLED_MetaCycleOff
        End If

        NotificationLED_GetSupportedFeatures = lngResult
    
    End If
    
    If Err.Number <> 0 Then 'Return 0 if error on API call.
        NotificationLED_GetSupportedFeatures = 0
    End If

End Function

Private Function NotificationLED_MakeNLED_SUPPORTS_INFO(ByVal LedNum As Long, _
                                                        ByVal lCycleAdjust As Long, _
                                                        ByVal fAdjustTotalCycleTime As Long, _
                                                        ByVal fAdjustOnTime As Long, _
                                                        ByVal fAdjustOffTime As Long, _
                                                        ByVal fMetaCycleOn As Long, _
                                                        ByVal fMetaCycleOff As Long) As String

    Dim varMembers As Variant

    varMembers = Array(UDTHelper_ToBinaryString(CLng(LedNum), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(lCycleAdjust), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(fAdjustTotalCycleTime), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(fAdjustOnTime), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(fAdjustOffTime), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(fMetaCycleOn), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(fMetaCycleOff), UDTHelper_CE_LONG))

    NotificationLED_MakeNLED_SUPPORTS_INFO = Join(varMembers, vbNullString)

End Function

Private Function NotificationLED_MakeNLED_SETTINGS_INFO(ByVal LedNum As Long, _
                                                        ByVal OffOnBlink As Long, _
                                                        ByVal TotalCycleTime As Long, _
                                                        ByVal OnTime As Long, _
                                                        ByVal OffTime As Long, _
                                                        ByVal MetaCycleOn As Long, _
                                                        ByVal MetaCycleOff As Long) As String

    Dim varMembers As Variant

    varMembers = Array(UDTHelper_ToBinaryString(CLng(LedNum), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(OffOnBlink), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(TotalCycleTime), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(OnTime), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(OffTime), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(MetaCycleOn), UDTHelper_CE_LONG), UDTHelper_ToBinaryString(CLng(MetaCycleOff), UDTHelper_CE_LONG))

    NotificationLED_MakeNLED_SETTINGS_INFO = Join(varMembers, vbNullString)

End Function



