VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登入"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "人員工號(8碼):"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()



    SetWindowPos Form2.hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub
