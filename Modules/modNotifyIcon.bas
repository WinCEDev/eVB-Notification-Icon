Attribute VB_Name = "NotifyIcon"
Option Explicit

'This module uses the UDT examples from the book "eMbedded Visual Basic: Windows CE and Pocket PC Mobile Applications".

Public Declare Function NotifyIcon_Shell_NotifyIcon _
               Lib "Coredll" _
               Alias "Shell_NotifyIcon" (ByVal dwMessage As Long, _
                                         ByVal pnid As String) As Long

Public Declare Function NotifyIcon_ImageList_GetIcon _
               Lib "Coredll" _
               Alias "ImageList_GetIcon" (ByVal himl As Long, _
                                          ByVal i As Long, _
                                          ByVal flags As Long) As Long

Public Declare Function NotifyIcon_DestroyIcon _
               Lib "Coredll" _
               Alias "DestroyIcon" (ByVal hIcon As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Binary String Datatype Sizes.

Private Const CE_INTEGER         As Long = 2

Private Const CE_LONG            As Long = 4

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NotifyIcon Messages.
'https://learn.microsoft.com/en-us/previous-versions/ms942613(v=msdn.10)

Private Const NIM_ADD            As Long = &H0 'Adds an icon to the status area.

Private Const NIM_MODIFY         As Long = &H1 'Modifies an icon in the status area.

Private Const NIM_DELETE         As Long = &H2 'Deletes an icon from the status area.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NotifyIcon Flags.
'https://learn.microsoft.com/en-us/previous-versions/ms961260(v=msdn.10)

Private Const NIF_MESSAGE        As Long = &H1 'The uCallbackMessage member is valid.

Private Const NIF_ICON           As Long = &H2 'The hIcon member is valid.

Private Const NIF_TIP            As Long = &H4 'The szTip member is valid.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Window Messages.

Public Const WM_LBUTTONDBLCLK    As Long = &H203 'Double-click.

Public Const WM_LBUTTONDOWN      As Long = &H201 'Button down.

Public Const WM_LBUTTONUP        As Long = &H202 'Button up.

Public Const WM_RBUTTONUP        As Long = &H205 'Button up.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Used by 'NotifyIcon_Add' and 'NotifyIcon_Remove'.

Private Const NOTIFYICONDATA_LEN As Long = 24 'Length in bytes of the NOTIFYICONDATA structure.

Private Const ICON_ID            As Long = 13 'Unique ID for this icon, since this module only lets you add a single icon, this value is hardcoded. Values 0-12 are reserved: https://learn.microsoft.com/en-us/previous-versions/windows/embedded/ms911889(v=msdn.10).

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Module-level variables.

Private CallbackFormHandle       As Long 'Holds the handle of the form receiving callbacks.

Private FormImageListHandle      As Long 'Holds the handle of the form's image list.

Private IconHandle               As Long 'Holds the handle of the icon used for the system tray.

Public Function NotifyIcon_Add(ByVal FormHandle As Long, _
                               ByVal ImageListHandle As Long, _
                               ByVal Index As Long) As Long
                               
    If CallbackFormHandle <> 0 Then 'An icon is already present.
        NotifyIcon_Remove 'First remove the old icon.
    End If

    CallbackFormHandle = FormHandle
    FormImageListHandle = ImageListHandle
    IconHandle = NotifyIcon_ImageList_GetIcon(FormImageListHandle, Index, 0)
    
    Dim strNOTIFYICONDATA As String
    
    strNOTIFYICONDATA = NotifyIcon_MakeNOTIFYICONDATA(NOTIFYICONDATA_LEN, CallbackFormHandle, ICON_ID, NIF_ICON Or NIF_MESSAGE, WM_LBUTTONDOWN, IconHandle)

    NotifyIcon_Add = NotifyIcon_Shell_NotifyIcon(NIM_ADD, strNOTIFYICONDATA)

End Function

Public Function NotifyIcon_Remove() As Long

    If CallbackFormHandle <> 0 Then

        Dim strNOTIFYICONDATA As String

        strNOTIFYICONDATA = NotifyIcon_MakeNOTIFYICONDATA(NOTIFYICONDATA_LEN, CallbackFormHandle, ICON_ID, 0, 0, 0)
    
        NotifyIcon_Remove = NotifyIcon_Shell_NotifyIcon(NIM_DELETE, strNOTIFYICONDATA) And NotifyIcon_DestroyIcon(IconHandle)

        CallbackFormHandle = 0
        FormImageListHandle = 0
        IconHandle = 0

    End If

End Function

Public Function NotifyIcon_Modify(ByVal Index As Long) As Long

    If CallbackFormHandle <> 0 Then

        Dim lngNewIcon As Long
        lngNewIcon = NotifyIcon_ImageList_GetIcon(FormImageListHandle, Index, 0) 'Load new icon.

        Dim strNOTIFYICONDATA As String

        strNOTIFYICONDATA = NotifyIcon_MakeNOTIFYICONDATA(NOTIFYICONDATA_LEN, CallbackFormHandle, ICON_ID, NIF_ICON Or NIF_MESSAGE, WM_LBUTTONDOWN, lngNewIcon)

        NotifyIcon_Modify = NotifyIcon_Shell_NotifyIcon(NIM_MODIFY, strNOTIFYICONDATA)
        
        'Delete old icon and load the new icon into the IconHandle variable.
        NotifyIcon_DestroyIcon IconHandle
        IconHandle = lngNewIcon

    End If

End Function

Private Function NotifyIcon_MakeNOTIFYICONDATA(ByVal cbSize As Long, _
                                               ByVal hWnd As Long, _
                                               ByVal uID As Long, _
                                               ByVal uFlags As Long, _
                                               ByVal uCallbackMessage As Long, _
                                               ByVal hIcon As Long) As String

    Dim varMembers As Variant

    varMembers = Array(NotifyIcon_ToBinaryString(CLng(cbSize), CE_LONG), NotifyIcon_ToBinaryString(CLng(hWnd), CE_LONG), NotifyIcon_ToBinaryString(CLng(uID), CE_LONG), NotifyIcon_ToBinaryString(CLng(uFlags), CE_LONG), NotifyIcon_ToBinaryString(CLng(uCallbackMessage), CE_LONG), NotifyIcon_ToBinaryString(CLng(hIcon), CE_LONG))

    NotifyIcon_MakeNOTIFYICONDATA = Join(varMembers, vbNullString)

End Function

Private Function NotifyIcon_GetByteValue(ByVal Number As Variant, _
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
    NotifyIcon_GetByteValue = Number And lngMask
    'Move that byte to the end of the number.
    NotifyIcon_GetByteValue = NotifyIcon_GetByteValue / (2 ^ (8 * BytePos))
    
End Function

Private Function NotifyIcon_ToBinaryString(ByVal Number As Variant, _
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
            NotifyIcon_ToBinaryString = NotifyIcon_ToBinaryString & (ChrB(NotifyIcon_GetByteValue(Number, i) + &H80))
        Else
            'Just append the byte to our string.
            NotifyIcon_ToBinaryString = NotifyIcon_ToBinaryString & ChrB(NotifyIcon_GetByteValue(Number, i))
        End If

    Next

End Function

