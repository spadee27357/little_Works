VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Macro Run�f���ܨt��"
   ClientHeight    =   13185
   ClientLeft      =   120
   ClientTop       =   1350
   ClientWidth     =   9030
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   13935
      Left            =   0
      ScaleHeight     =   13875
      ScaleWidth      =   8970
      TabIndex        =   0
      Top             =   0
      Width           =   9030
      Begin VB.CommandButton Command3 
         Caption         =   "�U�@��"
         Height          =   1095
         Left            =   2040
         TabIndex        =   3
         Top             =   10200
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�W�@��"
         Height          =   1095
         Left            =   240
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
         Width           =   8775
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   615
         Left            =   3840
         TabIndex        =   7
         Top             =   11040
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   495
         Left            =   5520
         TabIndex        =   6
         Top             =   10080
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   495
         Left            =   5160
         TabIndex        =   5
         Top             =   10560
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   495
         Left            =   3840
         TabIndex        =   4
         Top             =   10200
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   9855
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   8775
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Down As Integer
Public j As Integer
Public ActPLC As Object

Public ackPLC As Long
Public strConnPLC As String

Private bFirstIn As Boolean
Private nWidth, nHeight As Single
 Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
    End Type
     Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    'FindWindow ���o�����s��
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    'GetWindowRect ���o�����j�p
    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    'BringWindowToTop �N��������̤W�h
    Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long


Private Sub command2_Click()
   If Down > 1 Then
    Down = Down - 1
   ElseIf Down < 1 Then
    Down = Down
   ElseIf Down = 1 Then
    Down = Down
   End If

   Call Prdpicture(Down, g_sProductID)

End Sub

Private Sub Command3_Click()
        Dim nFile As String
        On Error Resume Next
    'C:\Documents and Settings\cfins03\�ୱ
    'Q:\CFCT
    'D:
    'ProductID = "FGF640XXQ1"
    'OperationID = "0190"
    
    j = 0
'--------------------------------D��
    nFile = Dir("D:\LogFile\MACRO RUN\" & ProductID & "\*.jpg")
    Do While Len(nFile)
        If nFile <> "" Then
            File = Split(nFile, "_")
            If File(1) <> "" Then 'Macro�q��
                j = j + 1
            End If
        End If
        nFile = Dir
    Loop
'--------------------------------D��
'--------------------------------�{�ɥ[�j��
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\*.txt")
    Do While Len(nFile) '�p���Ƨ������X�i�Ϥ�

        If nFile <> "" Then
            File = Split(nFile, "_")
            aFile = Split(File(1), ".")
            'Macro�q��
            If aFile(0) <> "" Then
                j = j + 1
            End If
        End If
        
        nFile = Dir
    Loop
'--------------------------------�{�ɥ[�j��
'--------------------------------�{�ɥ[�jALL��i

    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\ALL\*.jpg")
    Do While Len(nFile) '�p���Ƨ������X�i�Ϥ�

        If nFile <> "" Then
            j = j + 1
        End If
        nFile = Dir
    Loop
    nFile = ""

'--------------------------------�{�ɥ[�jALL��i
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\*.jpg")  '�]�w�Q�n�B�z���ؿ��� C:\123, �B�z���ɮװ��ɦW�� *.txt

    Do While Len(nFile)
        If nFile <> "" Then
            File = Split(nFile, "_")
            'Macro07����
            'If File(1) = "07" Or File(1) = "���p�I.jpg" Or File(1) = "����.jpg" Or File(1) = Left(OperationID, 2) Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then
            
            'Macro�q��
            If File(1) <> "" Then
                If File(1) = Left(OperationID, 2) Then
                    j = j + 1
                ElseIf CoaterID = "12" And File(1) = "R" Then
                    j = j + 1
                ElseIf CoaterID = "13" And File(1) = "L" Then
                    j = j + 1
                ElseIf File(1) <> "R" Then
                    If File(1) <> "L" Then
                        j = j + 1
                    End If
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

    Call Prdpicture(Down, g_sProductID)
'
End Sub


Private Sub MDIForm_Load()


'����~���{�� , �äW���ܲĤ@�h
 Dim i As Long, j As Long
 Dim rc As RECT
 i = FindWindow(vbNullString, "Macro Run�f���ܨt��")
 a = a
 GetWindowRect i, rc
 a = a
 BringWindowToTop i
'----------------------------------------
    Const SWP_NOMOVE = &H2 '����ʥثe������m
    Const SWP_NOSIZE = &H1 '����ʥثe�����j�p
    Const HWND_TOPMOST = -1 '-1�]�w���̤W�h  -2�����̤W�h
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos i, HWND_TOPMOST, 0, 0, 0, 1000, FLAGS

   ' Down = 1
   ' Call Prdpicture(Down, g_sProductID)
    Call localdata
 'Call Look
    Call ReadPLC

End Sub



Public Sub ReadPLC()



      On Error GoTo ErrorHandle


            a = a
            If exA_GlassID = GlassID Then
                GoTo Find
            End If
                Down = 1
                Call Prdpicture(Down, g_sProductID)
               
            
            'ex_GlassID = GlassID
            
            
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

            Form2.tmrScan.Interval = 120000  '500=�C0.5��j�M�@��
            'Form2.Timer1.Interval = 60000
      Exit Sub
ErrorHandle:
      'TraceOut "(ReadPLC) Err Line = " & Erl & " ,err.Description = " & err.Description, Error
      'Resume Next
End Sub


