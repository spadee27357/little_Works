VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�Z�ħY�ɨt��"
   ClientHeight    =   10725
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10725
   ScaleWidth      =   15585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   11400
      Top             =   9960
   End
   Begin VB.Timer Timer1 
      Left            =   10560
      Top             =   9960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�T�w"
      Height          =   735
      Left            =   6840
      TabIndex        =   0
      Top             =   9840
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   9615
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   15375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ref: http://www.programmer-club.com.tw/showSameTitleN/vb/17372.html(����{���̤p�ƪ����D)

Private Sub Command1_Click()
    Call Shell_NotifyIcon(NIM_ADD, nid)
    Me.Hide
End Sub

'��l�ѦҸ��:
'(1)http://www.hmhsieh.idv.tw/kjasp/ch16/Vb/RD/RUNPC/49.ASP#Q15 (���D15�G�N����Y�p�ɡA �Ʊ楦���ϥ���ܦb�u�@�C���k�U���C)
'(2)https://zhidao.baidu.com/question/45818243.html (vb6.0 �I���̤j�Ƴ̤p�ƩM�����]�k�W�����^Ĳ�o����ƥ�)
'(3)http://www.blueshop.com.tw/board/FUM200501271723350KG/BRD20080201170346EQD.html (�Ŧ�p�E - �Y�p��t�Τu��C)

'�{�����
Private Sub Form_Load()
    Dim FSO As FileSystemObject
    Dim str As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists("C:\LCD\RUN�f���I.txt") = False Then
        Comp = ""
    End If
    If FSO.FileExists("C:\LCD\RUN�f���I.txt") = True Then
        Open ("C:\LCD\RUN�f���I.txt") For Input As #1  '�}�Ҥ�r��
        Line Input #1, str              '�v��Ū��
            Comp = str
        Close #1
    End If
    ghWnd = Me.hWnd
    nid.cbSize = Len(nid)
    nid.hWnd = ghWnd
    nid.uID = uID
    nid.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
    nid.uCallbackMessage = WM_USER + 100
    nid.hIcon = Me.Icon
    nid.szTip = "�Z�ħY�ɨt��" & Chr(0)
    lngPrevWndProc = SetWindowLong(ghWnd, GWL_WNDPROC, AddressOf WndProc)
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3   '���̤W�h���
    Form1.Timer1.Interval = 60000 '60000
  
    
End Sub

'�{������
Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong ghWnd, GWL_WNDPROC, lngPrevWndProc
    Call Shell_NotifyIcon(NIM_DELETE, nid)
End Sub

Private Sub Form_Resize()
    Select Case Me.WindowState
        '�٭���s�Q���U�ε����j�p�o�ͧ���
        Case vbNormal
            
        '�̤p�ƫ��s�Q���U
        Case vbMinimized
            'Call Shell_NotifyIcon(NIM_ADD, nid)
            'Me.Hide
            
        '�̤j�ƫ��s�Q���U
        Case vbMaximized
        
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '������������^���ШD
    Select Case UnloadMode
        
        '�I�������������s
        'Cancel = true '�i�H�������������ʧ@
        Case vbFormControlMenu
            Call Shell_NotifyIcon(NIM_DELETE, nid)
            Cancel = True
            
        '�䥦�覡�ް_���������A�٦��䥦�`�ơA�o�̤��C�|
        Case Else
        
    End Select
End Sub



Private Sub Timer1_Timer()

    UTime = Format(Now, "hhmm")
       ' NTime = DateAdd("n", "2", UTime)
    'MsgBox "���I: " & Comp
    If Comp <> "" Then
        If UTime = "0715" Or UTime = "0915" Or UTime = "1115" Or UTime = "1315" Or UTime = "1515" Or UTime = "1715" _
           Or UTime = "1915" Or UTime = "2115" Or UTime = "2315" Or UTime = "0115" Or UTime = "0315" Or UTime = "0515" Then
            Call Shell_NotifyIcon(NIM_DELETE, nid)
            Form1.WindowState = vbNormal
            Form1.Show
            Form1.Command1.Visible = False
            If Comp = "All" Then
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\�s�y��\�ժ��M��\�C�B��\�Z�ĬݪO�Ϥ�\IMRV�ݪO.jpg")
                Delay (10)
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\�s�y��\�ժ��M��\�C�B��\�Z�ĬݪO�Ϥ�\INK�ݪO.jpg")
                Delay (10)
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\�s�y��\�ժ��M��\�C�B��\�Z�ĬݪO�Ϥ�\FMAC�ݪO.jpg")
                Delay (10)
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\�s�y��\�ժ��M��\�C�B��\�Z�ĬݪO�Ϥ�\RVRP�ݪO.jpg")
                Delay (10)
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\�s�y��\�ժ��M��\�C�B��\�Z�ĬݪO�Ϥ�\5-�`�Z�ĬݪO.jpg")
                Delay (5)
                Form1.Command1.Visible = True
            Else
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\�s�y��\�ժ��M��\�C�B��\�Z�ĬݪO�Ϥ�\" & Comp & "�ݪO.jpg")
                Delay (10)
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\�s�y��\�ժ��M��\�C�B��\�Z�ĬݪO�Ϥ�\5-�`�Z�ĬݪO.jpg")
                Delay (5)
                Form1.Command1.Visible = True
            End If
            'UTime = ""
        End If
    End If
End Sub

Private Sub Delay(ASecond As Integer)
    Dim before
    before = Timer
    Do
    DoEvents
    Loop Until (Int(Timer - before) = ASecond)
End Sub

