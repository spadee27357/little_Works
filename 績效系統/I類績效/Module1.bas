Attribute VB_Name = "Module1"
'�N�Ϥ���J�t�Χ�API
Public Comp As String
Public UTime As String
Public NTime As String
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_USER = &H400
Public Const uID = 9999
Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type
Public nid As NOTIFYICONDATA
Public Declare Function Shell_NotifyIcon Lib "Shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'����Windows Mouse
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const GWL_WNDPROC = (-4)
Public lngPrevWndProc As Long
Public ghWnd As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wParam = uID Then
    Select Case lParam
        Case WM_LBUTTONDOWN     '����@�U
            Call Shell_NotifyIcon(NIM_DELETE, nid)
            Form1.WindowState = vbNormal
            Form1.Show
        
        Case WM_LBUTTONDBLCLK   '�����U

        Case WM_RBUTTONDOWN     '�k��@�U
    '     �i�ܪ�椶�� �� �u�� �� �\���s�边�A�s�W�\���
    '     With FrmPop
    '     .PopupMenu .mnuTary, vbPopupMenuRightAlign, , , .mnuOpen
    '     End With
    End Select
End If
WndProc = CallWindowProc(lngPrevWndProc, hWnd, Msg, wParam, lParam)
End Function

