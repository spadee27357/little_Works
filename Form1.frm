VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "績效即時系統"
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
      Caption         =   "確定"
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
'Ref: http://www.programmer-club.com.tw/showSameTitleN/vb/17372.html(關於程式最小化的問題)

Private Sub Command1_Click()
    Call Shell_NotifyIcon(NIM_ADD, nid)
    Me.Hide
End Sub

'其餘參考資料:
'(1)http://www.hmhsieh.idv.tw/kjasp/ch16/Vb/RD/RUNPC/49.ASP#Q15 (問題15：將表單縮小時， 希望它的圖示顯示在工作列的右下角。)
'(2)https://zhidao.baidu.com/question/45818243.html (vb6.0 點擊最大化最小化和關閉（右上角的）觸發什麼事件)
'(3)http://www.blueshop.com.tw/board/FUM200501271723350KG/BRD20080201170346EQD.html (藍色小舖 - 縮小到系統工具列)

'程式初值
Private Sub Form_Load()
    Dim FSO As FileSystemObject
    Dim str As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists("C:\LCD\RUN貨站點.txt") = False Then
        Comp = ""
    End If
    If FSO.FileExists("C:\LCD\RUN貨站點.txt") = True Then
        Open ("C:\LCD\RUN貨站點.txt") For Input As #1  '開啟文字檔
        Line Input #1, str              '逐行讀取
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
    nid.szTip = "績效即時系統" & Chr(0)
    lngPrevWndProc = SetWindowLong(ghWnd, GWL_WNDPROC, AddressOf WndProc)
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3   '‘最上層顯示
    Form1.Timer1.Interval = 60000 '60000
  
    
End Sub

'程式結束
Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong ghWnd, GWL_WNDPROC, lngPrevWndProc
    Call Shell_NotifyIcon(NIM_DELETE, nid)
End Sub

Private Sub Form_Resize()
    Select Case Me.WindowState
        '還原按鈕被按下或視窗大小發生改變
        Case vbNormal
            
        '最小化按鈕被按下
        Case vbMinimized
            'Call Shell_NotifyIcon(NIM_ADD, nid)
            'Me.Hide
            
        '最大化按鈕被按下
        Case vbMaximized
        
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '捕獲到關閉窗体的請求
    Select Case UnloadMode
        
        '點擊視窗關閉按鈕
        'Cancel = true '可以取消關閉視窗動作
        Case vbFormControlMenu
            Call Shell_NotifyIcon(NIM_DELETE, nid)
            Cancel = True
            
        '其它方式引起視窗關閉，還有其它常數，這裡不列舉
        Case Else
        
    End Select
End Sub



Private Sub Timer1_Timer()

    UTime = Format(Now, "hhmm")
       ' NTime = DateAdd("n", "2", UTime)
    'MsgBox "站點: " & Comp
    If Comp <> "" Then
        If UTime = "0715" Or UTime = "0915" Or UTime = "1115" Or UTime = "1315" Or UTime = "1515" Or UTime = "1715" _
           Or UTime = "1915" Or UTime = "2115" Or UTime = "2315" Or UTime = "0115" Or UTime = "0315" Or UTime = "0515" Then
            Call Shell_NotifyIcon(NIM_DELETE, nid)
            Form1.WindowState = vbNormal
            Form1.Show
            Form1.Command1.Visible = False
            If Comp = "All" Then
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\績效看板圖片\IMRV看板.jpg")
                Delay (10)
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\績效看板圖片\INK看板.jpg")
                Delay (10)
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\績效看板圖片\FMAC看板.jpg")
                Delay (10)
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\績效看板圖片\RVRP看板.jpg")
                Delay (10)
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\績效看板圖片\5-總績效看板.jpg")
                Delay (5)
                Form1.Command1.Visible = True
            Else
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\績效看板圖片\" & Comp & "看板.jpg")
                Delay (10)
                Form1.Image1 = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\績效看板圖片\5-總績效看板.jpg")
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

