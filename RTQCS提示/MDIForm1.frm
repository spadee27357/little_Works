VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "RTQCS 提示系統"
   ClientHeight    =   14715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   14760
      Left            =   0
      ScaleHeight     =   14700
      ScaleWidth      =   18900
      TabIndex        =   0
      Top             =   0
      Width           =   18960
      Begin VB.CommandButton Command7 
         Height          =   855
         Left            =   22560
         Picture         =   "MDIForm1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   12120
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Height          =   855
         Left            =   10320
         Picture         =   "MDIForm1.frx":0993
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   12120
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Height          =   855
         Left            =   21480
         Picture         =   "MDIForm1.frx":1326
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   12120
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Height          =   855
         Left            =   9240
         Picture         =   "MDIForm1.frx":1CE7
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   12120
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "登入"
         Height          =   375
         Left            =   18480
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   14160
         TabIndex        =   19
         Text            =   "Text4"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "登入"
         Height          =   375
         Left            =   6120
         TabIndex        =   18
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   1335
         Left            =   12600
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "MDIForm1.frx":26A8
         Top             =   13200
         Width           =   12135
      End
      Begin VB.TextBox Text1 
         Height          =   1335
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "MDIForm1.frx":26AE
         Top             =   13200
         Width           =   12135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "登出"
         Height          =   375
         Left            =   18480
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "登出"
         Height          =   375
         Left            =   6120
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Label14"
         Height          =   375
         Left            =   15960
         TabIndex        =   22
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   375
         Left            =   3600
         TabIndex        =   21
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "站點:"
         Height          =   375
         Left            =   12840
         TabIndex        =   16
         Top             =   12360
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "產品:"
         Height          =   375
         Left            =   12840
         TabIndex        =   15
         Top             =   12000
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "產品:"
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   12000
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "人員工號(8碼):"
         Height          =   375
         Left            =   12960
         TabIndex        =   11
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Aligner:"
         Height          =   375
         Left            =   17160
         TabIndex        =   10
         Top             =   12360
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Coater:"
         Height          =   375
         Left            =   17160
         TabIndex        =   9
         Top             =   12000
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Gls ID:"
         Height          =   375
         Left            =   12840
         TabIndex        =   8
         Top             =   12720
         Width           =   4695
      End
      Begin VB.Label Label6 
         Caption         =   "人員工號(8碼):"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Aligner:"
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   12360
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Coater:"
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   12000
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Gls ID:"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   12720
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "站點:"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   12360
         Width           =   2535
      End
      Begin VB.Image Image2 
         Height          =   10815
         Left            =   12600
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   12135
      End
      Begin VB.Image Image1 
         Height          =   10815
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   12135
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long



Private Sub Command1_Click()
    MDIForm1.Text3 = ""
    Oper_1 = ""
    Product_1 = ""
    Operation_1 = ""
    Gls_1 = ""
    Coater_1 = ""
    Aligner_1 = ""
    MDIForm1.Label1 = "產品: "
    MDIForm1.Label2 = "站點: "
    MDIForm1.Label3 = "Gls_ID: "
    MDIForm1.Label4 = "Coater: "
    MDIForm1.Label5 = "Aligner: "
    MDIForm1.Label1.Font.Size = 13
    MDIForm1.Label2.Font.Size = 13
    MDIForm1.Label3.Font.Size = 13
    MDIForm1.Label4.Font.Size = 13
    MDIForm1.Label5.Font.Size = 13
    MDIForm1.Label13 = "尚未登入!"
    MDIForm1.Image1.Picture = LoadPicture("")
    MDIForm1.Text1 = ""
End Sub

Private Sub Command2_Click()
    MDIForm1.Text4 = ""
    Oper_2 = ""
    Product_2 = ""
    Operation_2 = ""
    Gls_2 = ""
    Coater_2 = ""
    Aligner_2 = ""
    MDIForm1.Label7 = "產品: "
    MDIForm1.Label8 = "站點: "
    MDIForm1.Label9 = "Gls_ID: "
    MDIForm1.Label10 = "Coater: "
    MDIForm1.Label11 = "Aligner: "
    MDIForm1.Label7.Font.Size = 13
    MDIForm1.Label8.Font.Size = 13
    MDIForm1.Label9.Font.Size = 13
    MDIForm1.Label10.Font.Size = 13
    MDIForm1.Label11.Font.Size = 13
    MDIForm1.Label14 = "尚未登入!"
    MDIForm1.Image2.Picture = LoadPicture("")
    MDIForm1.Text2 = ""
End Sub

Private Sub Command3_Click()
    If MDIForm1.Text3 <> "" Then
        Oper_1 = MDIForm1.Text3
        aFile = Dir("\\10.91.1.83\main$\RTQCS\" & Oper_1 & ".txt")
        If aFile = "" Then
            MDIForm1.Text3 = ""
            Oper_1 = ""
            MDIForm1.Label13 = "輸入錯誤!! 請重新登入!"
        Else
            MDIForm1.Image1.Picture = LoadPicture("")
            change1_1xxx = 9
            change1_2xxx = 9
            change1_5xxx = 9
            change1_7xxx = 9
            change1_8xxx = 9
            change1_DMRV = 9
            MDIForm1.Label13 = "人員 " & Oper_1 & " 已登入!"
            MDIForm1.Text1 = ""
        End If
    End If
End Sub

Private Sub Command4_Click()
    If MDIForm1.Text4 <> "" Then
        Oper_2 = MDIForm1.Text4
        aFile = Dir("\\10.91.1.83\main$\RTQCS\" & Oper_2 & ".txt")
        If aFile = "" Then
            MDIForm1.Text4 = ""
            Oper_2 = ""
            MDIForm1.Label14 = "輸入錯誤!! 請重新登入!"
        Else
            MDIForm1.Image2.Picture = LoadPicture("")
            change2_1xxx = 9
            change2_2xxx = 9
            change2_5xxx = 9
            change2_7xxx = 9
            change2_8xxx = 9
            change2_DMRV = 9
            MDIForm1.Label14 = "人員 " & Oper_2 & " 已登入!"
            MDIForm1.Text2 = ""
        End If
    End If
End Sub

Private Sub Command5_Click()
    Form1.Timer1.Interval = 1000
End Sub

Private Sub Command6_Click()
    Form1.Timer2.Interval = 1000
End Sub

Private Sub Command7_Click()
    Form1.Timer2.Interval = 0
End Sub

Private Sub Command9_Click()
    Form1.Timer1.Interval = 0
End Sub

Private Sub MDIForm_Load()
    MDIForm1.Text1 = ""
    MDIForm1.Text2 = ""
    MDIForm1.Text3 = ""
    MDIForm1.Text4 = ""
    MDIForm1.Label14 = "尚未登入!"
    MDIForm1.Label13 = "尚未登入!"
    Call localdata
    Call Prodpicture
    Form1.tmrScan.Interval = 1000
    
End Sub
