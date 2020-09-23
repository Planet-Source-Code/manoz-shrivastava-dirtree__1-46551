Attribute VB_Name = "modGlobals"
Public Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByVal cchBuf As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Public Function LvAutoSize(lv As ListView)
Dim Col2Adjst As Long, LngRc As Long
    For Col2Adjst = 0 To lv.ColumnHeaders.Count - 1
        LngRc = SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, Col2Adjst, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next
End Function

Public Function FormatSize(dwBytes As Single) As String
Dim sBuf As String
Dim dwBuf As Long
    sBuf = Space$(32)
    dwBuf = Len(sBuf)
    If StrFormatByteSize(dwBytes, sBuf, dwBuf) <> 0 Then
        FormatSize = Left$(sBuf, InStr(sBuf, Chr$(0)) - 1)
    End If
End Function
