Option Explicit
Option Base 1
Dim vFuncRegister As Variant
Dim vNative As Variant
 
Private Sub GetNative()
'
' get the list of native built-in functions and their attributes
'
If IsEmpty(vNative) Then
vNative = ThisWorkbook.Worksheets("NativeFuncs").Range("A2").Resize(ThisWorkbook.Worksheets("NativeFuncs").Range("A1000").End(xlUp).Row - 1, 3)
End If
End Sub
Private Sub GetFuncRegister()
'
' get list of registered functions and their funcIDs
'
Dim sCmd As String
Dim vRes As Variant
Dim j As Long
'
If Not IsEmpty(vFuncRegister) Then Exit Sub
'
vFuncRegister = Application.RegisteredFunctions     ''' get data on XLL registered functions
'
' add column for funcid
'
ReDim Preserve vFuncRegister(LBound(vFuncRegister) To UBound(vFuncRegister), 1 To 4) As Variant
'
For j = LBound(vFuncRegister) To UBound(vFuncRegister)
'
' get funcids
'
If vFuncRegister(j, 1) Like "*xll" Then
sCmd = "REGISTER.ID(""" & CStr(vFuncRegister(j, 1)) & """,""" & CStr(vFuncRegister(j, 2)) & """)"
vRes = Application.ExecuteExcel4Macro(sCmd)
If Not IsError(vRes) Then vFuncRegister(j, 4) = vRes
End If
Next j
End Sub

Private Function CheckFunc(strFunc As String, blMulti As Boolean, blVolatile As Variant) As String
'
' returns
' type = B if built-in, X if XLL else O for Other (VBA or Automation)
' blMulti is true if multithreaded
' BlVolatile is True if Volatile, False if not volatile and ? if don't know
'
Dim strType As String
Dim vFound As Variant
Dim j As Long
Dim strTypeString As String
Dim vExcelFuncID As Variant
'
blMulti = True
blVolatile = False
'
' check for native xl function
'
On Error Resume Next
vFound = Application.VLookup(strFunc, vNative, 1, False)
On Error GoTo 0
If Not IsError(vFound) Then
strType = "B"
If Application.VLookup(strFunc, vNative, 3, False) = "V" Then blVolatile = True
If Application.VLookup(strFunc, vNative, 2, False) = "S" Or Val(Application.Version) < 12 Then blMulti = False
End If
'
If Len(strType) = 0 Then
'
' get xlfuncid - if not error then its an XLL func
'
vExcelFuncID = Evaluate(strFunc)
If Not IsError(vExcelFuncID) Then
strType = "X"
For j = LBound(vFuncRegister) To UBound(vFuncRegister)
If strFunc = vFuncRegister(j, 2) Or vExcelFuncID = vFuncRegister(j, 4) Then
strTypeString = vFuncRegister(j, 3)
If InStr(strTypeString, "!") > 0 Or _
(InStr(strTypeString, "#") > 0 And (InStr(strTypeString, "R") > 0 Or InStr(strTypeString, "U") > 0)) _
Then blVolatile = True
If InStr(strTypeString, "$") = 0 Or Val(Application.Version) < 12 Then blMulti = False
Exit For
End If
Next j
End If
End If
'
If Len(strType) = 0 Then
'
' else its Other (VBA or Automation)
'
strType = "O"
blMulti = False     ''' cant be multi
blVolatile = "?"    ''' don't know if volatile
End If
'
CheckFunc = strType
End Function

Public Function IsMultiThreaded(strFuncName As String) As Variant
'
' check if a function is Multi-Threaded
' Returns true or false
'
Dim blMulti As Boolean
Dim blVolatile As Variant
Dim strType As String
'
GetNative
GetFuncRegister
strType = CheckFunc(strFuncName, blMulti, blVolatile)
'
IsMultiThreaded = blMulti
End Function
Public Function IsVolatile(strFuncName As String) As Variant
'
' check if a function is volatile
' returns True or False or ? if don't know
'
Dim blMulti As Boolean
Dim blVolatile As Variant
Dim strType As String
'
GetNative
GetFuncRegister
strType = CheckFunc(strFuncName, blMulti, blVolatile)
'
IsVolatile = blVolatile
End Function
Public Function FuncType(strFuncName As String) As Variant
'
' get type of function: B for built-in Excel, X for XLL, O for Other
 
Dim blMulti As Boolean
Dim blVolatile As Variant
Dim strType As String
'
GetNative
GetFuncRegister
FuncType = CheckFunc(strFuncName, blMulti, blVolatile)
End Function
