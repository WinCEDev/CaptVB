Attribute VB_Name = "UDTHelper"
Option Explicit

'This module uses the UDT examples from the book "eMbedded Visual Basic: Windows CE and Pocket PC Mobile Applications".

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Binary String Datatype Sizes.

Public Const UDTHelper_CE_INTEGER As Long = 2

Public Const UDTHelper_CE_LONG    As Long = 4

Private Function UDTHelper_GetByteValue(ByVal Number As Variant, _
                                        ByVal BytePos As Integer) As Long
    
    Dim lngMask As Long

    On Error Resume Next

    'Cannot check byte positions other than 0 to 3.
    If BytePos > 3 Or BytePos < 0 Then

        Exit Function

    End If

    If BytePos < 3 Then
        'Build a lngMask of all bits on for the desired byte.
        lngMask = &HFF * (2 ^ (8 * BytePos))
    Else
        'The last bit is reserved for sign (+/-).
        lngMask = &H7F * (2 ^ (8 * BytePos))
    End If

    'Turn off all bits but the byte we're after.
    UDTHelper_GetByteValue = Number And lngMask
    'Move that byte to the end of the number.
    UDTHelper_GetByteValue = UDTHelper_GetByteValue / (2 ^ (8 * BytePos))
    
End Function

Private Function UDTHelper_ToBinaryString(ByVal Number As Variant, _
                                          ByVal Bytes As Integer) As String

    Dim blnIsNegative As Boolean

    'Cannot check byte positions other than 0 to 3.
    If Bytes > 4 Or Bytes < 1 Then

        Exit Function

    End If

    'If the number is negative, we need to handle it last, so we'll set a flag.
    If Number < 0 Then
        blnIsNegative = True

        'Get the absolute value.
        Number = Number * -1

        'Get the binary complement (except the most sign. bit).
        Number = Number Xor ((2 ^ (8 * Bytes - 1)) - 1)

        'Add one.
        Number = Number + 1
    End If
    
    Dim i As Long

    'Start at the least significant bit (0) and work backwards.
    For i = 0 To Bytes - 1

        If i = Bytes - 1 And blnIsNegative Then
            'If the number is negative we must turn on the most significant bit and then and append it to the string.
            UDTHelper_ToBinaryString = UDTHelper_ToBinaryString & (ChrB(UDTHelper_GetByteValue(Number, i) + &H80))
        Else
            'Just append the byte to our string.
            UDTHelper_ToBinaryString = UDTHelper_ToBinaryString & ChrB(UDTHelper_GetByteValue(Number, i))
        End If

    Next

End Function

Public Function UDTHelper_FromBinaryString(BString As String) As Variant

    Dim i           As Integer

    Dim bIsNegative As Boolean

    bIsNegative = False

    ' start at the end of the string and work backwards
    For i = LenB(BString) - 1 To 0 Step -1

        If i = LenB(BString) - 1 And (AscB(MidB(BString, i + 1, 2)) And &H80) Then
            ' check the signigicant bit
            ' If it's negative, set a flag
            bIsNegative = True
        End If

        If bIsNegative = True Then
            ' extract the binary complement of the byte
            UDTHelper_FromBinaryString = UDTHelper_FromBinaryString + ((AscB(MidB(BString, i + 1, 2)) Xor &HFF) * (2 ^ (8 * i)))
        Else
            ' extract the byte
            UDTHelper_FromBinaryString = UDTHelper_FromBinaryString + AscB(MidB(BString, i + 1, 2)) * (2 ^ (8 * i))
        End If

    Next

    ' if it is supposed to be negative, make it so
    If bIsNegative Then
        ' Subtract one
        UDTHelper_FromBinaryString = UDTHelper_FromBinaryString + 1

        ' make it negative
        UDTHelper_FromBinaryString = UDTHelper_FromBinaryString * -1
    End If

End Function



