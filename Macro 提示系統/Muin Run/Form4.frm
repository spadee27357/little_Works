VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   11550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9420
   LinkTopic       =   "Form4"
   ScaleHeight     =   13755
   ScaleMode       =   0  'User
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "離開"
      Height          =   1215
      Left            =   7320
      TabIndex        =   2
      Top             =   10080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一頁"
      Height          =   1215
      Left            =   2280
      TabIndex        =   1
      Top             =   10080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "上一頁"
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   10080
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   9735
      Left            =   120
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()

    Form1.Text1 = ""
    Form4.Hide
    MDIForm1.Show
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '捕獲到關閉窗体的請求
    Select Case UnloadMode
        
        '點擊視窗關閉按鈕
        Case vbFormControlMenu
            Form3.Show
            
        '其它方式引起視窗關閉，還有其它常數，這裡不列舉
        Case Else
        
    End Select
End Sub

