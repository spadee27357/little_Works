VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Macro Run貨提示系統"
   ClientHeight    =   13230
   ClientLeft      =   120
   ClientTop       =   1350
   ClientWidth     =   11115
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   13935
      Left            =   0
      ScaleHeight     =   13875
      ScaleWidth      =   11055
      TabIndex        =   0
      Top             =   0
      Width           =   11115
      Begin VB.CommandButton Command3 
         Caption         =   "下一頁"
         Height          =   1095
         Left            =   2400
         TabIndex        =   3
         Top             =   10200
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "上一頁"
         Height          =   1095
         Left            =   600
         TabIndex        =   2
         Top             =   10200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   11520
         Width           =   10815
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   615
         Left            =   5400
         TabIndex        =   7
         Top             =   11040
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   495
         Left            =   7680
         TabIndex        =   6
         Top             =   10080
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   495
         Left            =   7200
         TabIndex        =   5
         Top             =   10560
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   495
         Left            =   4440
         TabIndex        =   4
         Top             =   10800
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   9855
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   10815
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public j As Integer

Private bFirstIn As Boolean
Private nWidth, nHeight As Single
 Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
    End Type
     Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    'FindWindow 取得視窗編號
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    'GetWindowRect 取得視窗大小
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    'BringWindowToTop 將視窗移到最上層
    Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long


Private Sub command2_Click()
   If Down > 1 Then
    Down = Down - 1
   ElseIf Down < 1 Then
    Down = Down
   ElseIf Down = 1 Then
    Down = Down
   End If

   Call Prdpicture

End Sub

Private Sub Command3_Click()
        Dim FSO As FileSystemObject
        Dim nFile As String
        On Error Resume Next
    'C:\Documents and Settings\cfins03\桌面
    'Q:\CFCT
    'D:
    ProductID = "FGF654XFQ1"
    'OperationID = "0190"
    
    j = 0
    
'    nFile = Dir("D:\LogFile\MACRO RUN\" & ProductID & "\*.jpg")
    
'    Do While Len(nFile)
'        If nFile <> "" Then
'            File = Split(nFile, "_")
'            'If File(2) = "" Or File(1) = Left(OperationID, 2) Or File(1) = "07" Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then  'Macro07條件
'            If File(2) = "" Or File(1) = Left(OperationID, 2) Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then 'Macro通用
'                j = j + 1
'            End If
'        End If
'        nFile = Dir
'    Loop
    
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\Final Macro圖片\*.txt")  '設定想要處理的目錄為 C:\123, 處理的檔案副檔名為 *.txt

    Do While Len(nFile)
        If nFile <> "" Then
            File = Split(nFile, "_")
            aFile = Split(nFile, ".")
            Set FSO = CreateObject("Scripting.FileSystemObject")
            If FSO.FileExists("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\Final Macro圖片\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".jpg") Then
           
                If File(1) <> "" Then
                    j = j + 1
                End If
            End If
        End If
        nFile = Dir
    Loop
    
    If Down < j Then
        Down = Down + 1
    ElseIf Down > j Then
        Down = j
    ElseIf Down = j Then
        Down = j
    End If

    Call Prdpicture
'
End Sub


Private Sub MDIForm_Load()


' MDIForm1.WindowState = 0
'Form2.tmrScan.Interval = 500
'----------------------------------------------------------------
'----------------------------------------------------------------
'抓取外部程式 , 並上移至第一層
 Dim i As Long
 Dim rc As RECT
 i = FindWindow(vbNullString, "Macro Run貨提示系統")
 a = a
 GetWindowRect i, rc
 a = a
 BringWindowToTop i
'----------------------------------------
    Const SWP_NOMOVE = &H2 '不更動目前視窗位置
    Const SWP_NOSIZE = &H1 '不更動目前視窗大小
    Const HWND_TOPMOST = -1 '-1設定為最上層  -2取消最上層
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos i, HWND_TOPMOST, 0, 0, 0, 1000, FLAGS

   ' Down = 1
   ' Call Prdpicture()
    Call localdata
    Call Read

End Sub
Public Sub Read()

      Dim i As Integer
      On Error GoTo ErrorHandle

            If exA_GlassID = GlassID Then
                GoTo Find
            End If
                Down = 1
                Call Prdpicture
  
Find:
            If GlassID = "" Then
                MDIForm1.Image1.Picture = LoadPicture("")
                MDIForm1.Label1 = ""
                MDIForm1.Label2 = "Product ID: "
                MDIForm1.Label3 = "Operation ID:  "
                MDIForm1.Label4 = ""
                MDIForm1.Text1.Text = ""
                MDIForm1.Label1.Font.Size = 15
                MDIForm1.Label2.Font.Size = 15
                MDIForm1.Label3.Font.Size = 15
                MDIForm1.Label4.Font.Size = 15
            End If

            Form2.tmrScan.Interval = 10000  '500=每0.5秒搜尋一次

      Exit Sub
ErrorHandle:

End Sub

