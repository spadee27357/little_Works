VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3525
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1785
   ScaleWidth      =   3525
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer2 
      Left            =   2640
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   720
   End
   Begin VB.Timer tmrScan 
      Left            =   1200
      Top             =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    a = a
    If change1_1xxx > 8 Then
        change1_1xxx = change1_1xxx - 1
    End If
    If change1_2xxx > 8 Then
        change1_2xxx = change1_2xxx - 1
    End If
    If change1_5xxx > 8 Then
        change1_5xxx = change1_5xxx - 1
    End If
    If change1_7xxx > 8 Then
        change1_7xxx = change1_7xxx - 1
    End If
    If change1_8xxx > 8 Then
        change1_8xxx = change1_8xxx - 1
    End If
    If change1_DMRV > 8 Then
        change1_DMRV = change1_DMRV - 1
    End If

End Sub

Private Sub Timer2_Timer()

    If change2_1xxx > 8 Then
        change2_1xxx = change2_1xxx - 1
    End If
    If change2_2xxx > 8 Then
        change2_2xxx = change2_2xxx - 1
    End If
    If change2_5xxx > 8 Then
        change2_5xxx = change2_5xxx - 1
    End If
    If change2_7xxx > 8 Then
        change2_7xxx = change2_7xxx - 1
    End If
    If change2_8xxx > 8 Then
        change2_8xxx = change2_8xxx - 1
    End If
    If change2_DMRV > 8 Then
        change2_DMRV = change2_DMRV - 1
    End If
End Sub

Private Sub tmrScan_Timer()
     Call localdata

     If ex_Product_1 <> Product_1 And ex_Product_1 <> "" Then
        change1_1xxx = 9
        change1_2xxx = 9
        change1_5xxx = 9
        change1_7xxx = 9
        change1_8xxx = 9
        change1_DMRV = 9
        Form1.Timer1.Interval = 0
     End If
     
     If ex_Product_2 <> Product_2 And ex_Product_2 <> "" Then
        change2_1xxx = 9
        change2_2xxx = 9
        change2_5xxx = 9
        change2_7xxx = 9
        change2_8xxx = 9
        change2_DMRV = 9
        Form1.Timer2.Interval = 0
     End If
     change = change + 1
     Call Prodpicture
     
End Sub
