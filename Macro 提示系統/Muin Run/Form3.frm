VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   1995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4935
   LinkTopic       =   "Form3"
   ScaleHeight     =   1995
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CNVR As String
Private Sub Command1_Click()
    CNVR = Combo1.Text
    Form3.Hide
    Form4.Show
End Sub

Private Sub Form_Load()

Combo1.Clear
Combo1.AddItem "CNVR01"
Combo1.AddItem "CNVR02"
Combo1.AddItem "CNVR03"
Combo1.AddItem "CNVR04"
Combo1.AddItem "CNVR05"
Combo1.AddItem "CNVR06"
Combo1.AddItem "CNVR07"
Combo1.AddItem "CNVR08"
Combo1.AddItem "CNVR09"
Combo1.AddItem "CNVR10"
Combo1.AddItem "CNVR11"
Combo1.AddItem "CNVR12"
Combo1.AddItem "CNVR13"
Combo1.AddItem "CNVR14"
Combo1.AddItem "CNVR15"
Combo1.AddItem "CNVR16"
Combo1.AddItem "CNVR17"
Combo1.AddItem "CNVR18"
Combo1.AddItem "CNVR19"
Combo1.AddItem "CNVR20"


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '捕獲到關閉窗体的請求
    Select Case UnloadMode
        
        '點擊視窗關閉按鈕
        Case vbFormControlMenu
            MDIForm1.Show
            
        '其它方式引起視窗關閉，還有其它常數，這裡不列舉
        Case Else
        
    End Select
End Sub

