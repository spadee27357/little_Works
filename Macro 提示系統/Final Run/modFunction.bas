Attribute VB_Name = "modFunction"
'Creat by herbert @ 2010.09.02

Option Explicit
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long        '<<52.14.731 dp87101 ¼W¥[Messagebox function
Private Const RECIPE_BODY_NAME As String = "RecipeBody.ini"

Public Sub Check_Folder(Folder As String)
1        Dim FSO As Object
2        If Dir(Folder, vbDirectory) = "" Then
3            Set FSO = CreateObject("Scripting.FileSystemObject")
4            FSO.CreateFolder Folder
5            Set FSO = Nothing
6        End If
End Sub

Public Function GetIniString(sSec As String, sKey As String, sFile As String) As String
1        Dim sBuf As String
2        Dim nSize As Integer
3        Dim nTemp As Long
4        Dim sTemp As String
   
6        nSize = 100
7        sBuf = Space(nSize)
   
9        nTemp = GetPrivateProfileString(sSec, sKey, "", sBuf, nSize, sFile)
10       sTemp = RemoveSpace(sBuf)
   
12       GetIniString = sTemp
End Function

Public Function FileExists(sFileName As String) As Boolean
1        Dim FSO As New Scripting.FileSystemObject
2        FileExists = FSO.FileExists(sFileName)
End Function

Public Function WriteIniString(ByVal FileName As String, ByVal Sec As String, ByVal Key As String, ByVal Datas As String) As Boolean
1        Dim sDummy As Long
   
3        sDummy = -99
4        On Error GoTo ErrWriteIniString

6        sDummy = WritePrivateProfileString(Sec, Key, Datas, FileName)
7        If sDummy >= 0 Then WriteIniString = True
8        Exit Function

10 ErrWriteIniString:
11       WriteIniString = False
12       Exit Function
End Function


Public Function RemoveSpace(sValue As String) As String
1        Dim sTemp As String
2        Dim i As Integer
   
4        For i = 1 To Len(sValue)
5            If Mid(sValue, i + 1, 2) <> "  " Then
6                sTemp = sTemp & Mid(sValue, i, 1)
7            Else
8                GoTo break
9            End If
10       Next

12 break:
13       RemoveSpace = sTemp
End Function

