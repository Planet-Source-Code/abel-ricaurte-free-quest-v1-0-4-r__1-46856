Attribute VB_Name = "mdlRewrite"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public varTmp As Variant, strFree As String, strFile As String, e%, i%

Public Function ReadINI(Section, KeyName, filename As String) As String
  Dim sRet As String
  sRet = String(998, Chr(0))
  ReadINI = Left(sRet, GetPrivateProfileString(Section, KeyName, "", sRet, Len(sRet), filename))
End Function

Public Function WriteINI(Section, KeyName, NewString As String, filename As String) As String
  Dim sWet As String
  sWet = WritePrivateProfileString(Section, KeyName, NewString, filename)
End Function

Public Function PStr(strCase As String, strFull As String, Optional strText As String, Optional strKey As String, Optional Index As Double) As Variant
Dim strRpl As String, strCur As String, strSub1 As String, strSub2 As String, iCur As Integer, iText As Integer, iKey As Integer
  iText = IIf(strText = "", 0, InStr(strFull & "|", "" & strText & "|")): iKey = IIf(strKey = "", 0, InStr(strFull, "|" & strKey & "")): iCur = IIf(iKey > 0, iKey, iText)
  If strCase = "Add" Then strRpl = "|" & strKey & "" & strText: strFull = strFull & IIf(iKey > 0, "", strRpl): iCur = InStr(strFull, "|" & strKey & "")
  If iCur > 0 Then Index = UBound(Split(Left(strFull, iCur), "|")) Else Index = IIf(Index <= UBound(Split(strFull, "|")), Index, 0)
  If Index > 0 Then strCur = "|" & Split(strFull, "|")(Index): strSub1 = Mid(Split(strCur, "")(0), 2): strSub2 = Split(strCur, "")(1)
  If strCase = "Find" Then PStr = IIf(strKey = "Return", strSub1, IIf(strText = "Return", strSub2, Index)) Else PStr = Replace(strFull, strCur, strRpl)
End Function

Public Sub SplitAndAdd(objBox As Object, varAdd As Variant)
  Dim varOpt As Variant
  varOpt = Split(varAdd, "")
  For i% = 1 To UBound(varOpt)
    objBox.AddItem varOpt(i%)
  Next i%
End Sub

Public Sub ObjToObj(objDest As Object, objTmp As Object, Optional strSkip As String)
  For i% = 0 To objTmp.ListCount - 1
    If strSkip <> objTmp.List(i%) Then objDest.AddItem objTmp.List(i%)
  Next i%
End Sub

Public Function GetIndex(objBox As Object, strFind As String, Optional strCase As String) As Integer
  For i% = 0 To objBox.ListCount - 1
    If LCase(objBox.List(i%)) = LCase(strFind) Then If strCase = "Remove" Then objBox.RemoveItem i% Else GetIndex = i%
  Next i%
End Function
