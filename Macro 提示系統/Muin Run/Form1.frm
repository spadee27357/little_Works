VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   4995
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "�T�w"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Pass Word:"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If Form1.Text1 = "ENG" Or Form1.Text1 = "eng" Then
        Form1.Hide
        Form3.Show
    Else
        Form1.Text1 = ""
        Form1.Label2 = "��J���~!!"
    End If
    
End Sub

Private Sub Form_Load()
    Form1.Label2 = ""
    Form1.Text1 = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '������������^���ШD
    Select Case UnloadMode
        
        '�I�������������s
        Case vbFormControlMenu
            MDIForm1.Show
            
        '�䥦�覡�ް_���������A�٦��䥦�`�ơA�o�̤��C�|
        Case Else
        
    End Select
End Sub

