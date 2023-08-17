Attribute VB_Name = "PathHelpers"
Option Explicit

Public Declare Function PathHelpers_GetTempPathW _
               Lib "Coredll" _
               Alias "GetTempPathW" (ByVal nBufferLength As Long, _
                                     ByVal lpBuffer As String) As Long

Public Const PathHelpers_MAX_PATH = 260

Public Function PathHelpers_RemoveExtension(ByVal FilePath As String) As String

    PathHelpers_RemoveExtension = Left(FilePath, InStrRev(FilePath, ".") - 1)

End Function

Public Function PathHelpers_AddPathSeparator(ByVal FilePath As String) As String

    If LenB(FilePath) <> 0 Then
        If FilePath <> "\" Then 'Ignore root path.
            If Right(FilePath, 1) <> "\" Then
                PathHelpers_AddPathSeparator = FilePath & "\"
            Else
                PathHelpers_AddPathSeparator = FilePath
            End If

        Else
            PathHelpers_AddPathSeparator = FilePath
        End If
    End If

End Function

Public Function PathHelpers_RemovePathSeparator(ByVal FilePath As String) As String

    If LenB(FilePath) <> 0 Then
        If FilePath <> "\" Then 'Ignore root path.
            If Right(FilePath, 1) = "\" Then
                PathHelpers_RemovePathSeparator = Left(FilePath, Len(FilePath) - 1)
            Else
                PathHelpers_RemovePathSeparator = FilePath
            End If

        Else
            PathHelpers_RemovePathSeparator = FilePath
        End If
    End If
    
End Function

Public Function PathHelpers_GetTempPath() As String

    Dim strTempPath As String

    strTempPath = String(PathHelpers_MAX_PATH, vbNullChar)

    Dim lngLength As Long

    lngLength = PathHelpers_GetTempPathW(LenB(strTempPath), strTempPath)

    strTempPath = Left(strTempPath, lngLength)

    PathHelpers_GetTempPath = strTempPath

End Function

Public Function PathHelpers_GetTempFileName(ByVal PathName As String, _
                                            ByVal PrefixString As String, _
                                            ByVal Unique As Integer) As String

    If Unique = 0 Then
        Randomize
        Unique = CInt(Rnd * 32768)
    End If

    PathHelpers_GetTempFileName = PathHelpers_AddPathSeparator(PathName) & Left(PrefixString, 3) & Hex(Unique) & ".TMP"

End Function

Public Function PathHelpers_ContainsInvalidChars(ByVal FilePath As String, _
                                                 ByVal IsFileName As Boolean) As Boolean

    Dim varInvalidChars As Variant

    varInvalidChars = Array("/", ":", "*", "?", """", "<", ">", "|")
    
    If IsFileName Then

        Dim lngUBound As Long

        lngUBound = UBound(varInvalidChars) + 1
    
        ReDim Preserve varInvalidChars(lngUBound)
        varInvalidChars(lngUBound) = "\"
    End If

    Dim Char As String

    For Each Char In varInvalidChars

        If InStr(FilePath, Char) <> 0 Then
            PathHelpers_ContainsInvalidChars = True

            Exit Function

        End If

    Next

    PathHelpers_ContainsInvalidChars = False

End Function

Public Function PathHelpers_GetFilenameFromPath(ByVal FilePath As String) As String
    PathHelpers_GetFilenameFromPath = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "\"))
End Function

Public Function PathHelpers_GetNextAvailableFileName(ByRef FileSystem As FileSystem, _
                                                     ByVal BasePath As String, _
                                                     ByVal FileName As String) As String

    Dim strFilePath      As String

    Dim strFileExtension As String

    Dim strBaseName      As String

    Dim lngCounter       As Long
    
    ' Split the filename and extension
    strFileExtension = Right(FileName, Len(FileName) - InStrRev(FileName, ".") + 1)
    strBaseName = Left(FileName, Len(FileName) - Len(strFileExtension))

    ' Start with no counter
    lngCounter = 0
    
    Do

        ' Construct the full file path to check if the file exists
        If lngCounter = 0 Then
            strFilePath = BasePath & strBaseName & strFileExtension
        Else
            strFilePath = BasePath & strBaseName & " (" & lngCounter & ")" & strFileExtension
        End If
        
        ' Check if the file exists
        If Len(FileSystem.Dir(strFilePath)) = 0 Then
            ' File doesn't exist, return this filename as the next available one
            PathHelpers_GetNextAvailableFileName = strFilePath

            Exit Function

        End If
        
        ' Increment the counter for the next iteration
        lngCounter = lngCounter + 1
    Loop

End Function



