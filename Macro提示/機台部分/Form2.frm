VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   LinkTopic       =   "MDIForm1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrScan 
      Interval        =   10000
      Left            =   720
      Top             =   720
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Net/h Card
Private Declare Function mdopen Lib "MDFUNC32.DLL" (ByVal Chan As Integer, ByVal Mode As Integer, Path As Long) As Integer
Private Declare Function mdclose Lib "MDFUNC32.DLL" (ByVal Path As Long) As Integer
Private Declare Function mdreceive Lib "MDFUNC32.DLL" (ByVal Path As Long, ByVal Stno As Integer, ByVal Devtyp As Integer, ByVal Devno As Integer, Size As Integer, Buf As Any) As Integer


Private Sub Form_Load()
'add
End Sub



Private Sub tmrScan_Timer()

     On Error GoTo ErrorHandle
    'MsgBox "timer success"
     'Get RecipeNo, GroupNo, ProductID, GroupID, GlassID
     
     Call localdata
     Call MDIForm1.ReadPLC

     Exit Sub
ErrorHandle:
     
End Sub







