Attribute VB_Name = "PathHelpers"
Option Explicit

Public Function PathHelpers_RemoveExtension(ByVal FilePath As String) As String

    PathHelpers_RemoveExtension = Left(FilePath, InStrRev(FilePath, ".") - 1)

End Function

Public Function PathHelpers_AddPathSeparator(ByVal FilePath As String) As String

    If LenB(FilePath) <> 0 Then
        If Right(FilePath, 1) <> "\" Then
            PathHelpers_AddPathSeparator = FilePath & "\"
        End If
    End If

End Function


