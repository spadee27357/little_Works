Attribute VB_Name = "�Ϥ�"
Private bFirstIn As Boolean
Private nWidth, nHeight As Single


Sub Prdpicture()

    On Error Resume Next
    Dim nFile As String
    Dim str As String
    Dim b As String
    Dim message() As String
    Dim FSO As FileSystemObject

    ProductID = "FGF654XFQ1"
    'OperationID = "5690"
    'Down = 2
    a = a
    MDIForm1.WindowState = 0
    
    If ProductID = "" Then
        MDIForm1.Image1.Picture = LoadPicture("")
        MDIForm1.Label1 = ""
        MDIForm1.Label2 = "Product ID: "
        MDIForm1.Label3 = "Operation ID: "
        MDIForm1.Label2.Font.Size = 15
        MDIForm1.Label3.Font.Size = 15
        MDIForm1.Label4 = ""
        MDIForm1.Text1.Text = ""
        GoTo finish
    End If

    

    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\Final Macro�Ϥ�\*.txt")  '�]�w�Q�n�B�z���ؿ��� C:\123, �B�z���ɮװ��ɦW�� *.txt

    If nFile = "" Then  '�n�O�S���ɮ� �ǵL����
        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\�L����\�L����.jpg")
        MDIForm1.Label2 = "ProductID:  " & ProductID
        MDIForm1.Label3 = "Operation ID:  " & OperationID
        MDIForm1.Label4 = ""
        MDIForm1.Text1.Text = ""
        MDIForm1.Label1.Font.Size = 15
        MDIForm1.Label2.Font.Size = 15
        MDIForm1.Label3.Font.Size = 15
        MDIForm1.Label4.Font.Size = 15
        Down = 0
    End If

    nFile = ""
'--------------------------------D��
'    nFile = Dir("D:\LogFile\MACRO RUN\" & ProductID & "\*.jpg")
'        Do While Len(nFile)
'    If nFile <> "" Then
'            File = Split(nFile, "_")
'            'If File(2) = "" Or File(1) = Left(OperationID, 2) Or File(1) = "07" Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then 'Macro07����
'            If File(2) = "" Or File(1) = Left(OperationID, 2) Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then 'Macro�q��
'            'If CoaterID = "12" And File(1) = "R" Then  'MUIN�q��
'                j = j + 1
'                If j = Down Then
'                    MDIForm1.Image1.Picture = LoadPicture("D:\LogFile\MACRO RUN\" & ProductID & "\" & nFile)
'                    xFile = Split(nFile, ".")
'                    MDIForm1.Label4 = xFile(0)
'                    MDIForm1.Label2 = "ProductID:  " & ProductID
'                    MDIForm1.Label3 = "Operation ID:  " & OperationID
'                    MDIForm1.Label1.Font.Size = 15
'                    MDIForm1.Label2.Font.Size = 15
'                    MDIForm1.Label3.Font.Size = 15
'                    MDIForm1.Label4.Font.Size = 15
'                End If
'            End If
'    End If
'    nFile = Dir
'    Loop
'--------------------------------D��
    
    
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\Final Macro�Ϥ�\*.txt")
    Do While Len(nFile) '�p���Ƨ������X�i�Ϥ�
        If nFile <> "" Then
            File = Split(nFile, "_")
            aFile = Split(nFile, ".")
            Set FSO = CreateObject("Scripting.FileSystemObject")
            If FSO.FileExists("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\Final Macro�Ϥ�\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".jpg") Then
                If File(1) <> "" Then
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\Final Macro�Ϥ�\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".jpg")
                        xFile = Split(nFile, ".")
                        MDIForm1.Label4 = xFile(0)
                        MDIForm1.Label2 = "ProductID:  " & ProductID
                        MDIForm1.Label3 = "Operation ID:  " & OperationID
                        MDIForm1.Label1.Font.Size = 15
                        MDIForm1.Label2.Font.Size = 15
                        MDIForm1.Label3.Font.Size = 15
                        MDIForm1.Label4.Font.Size = 15
                    End If
                End If

            End If
            
        End If
        
        nFile = Dir
    Loop
    j = 0

'--------------------------------D��
'     Do While Len(nFile)
'nFile = Dir("D:\LogFile\MACRO RUN\" & ProductID & "\*.txt")
'
'
'        If nFile <> "" Then
'            File = Split(nFile, "_")
'            'If File(2) = "" Or File(1) = Left(OperationID, 2) Or File(1) = "07" Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then  'Macro07����
'            If File(2) = "" Or File(1) = Left(OperationID, 2) Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then 'Macro�q��
'            a = a
'            'If CoaterID = "12" And File(1) = "R" Then     'MUIN�q��
'                j = j + 1
'                If j = Down Then
'                    MDIForm1.Text1.Text = ""             '�M�����e
'                    Open ("D:\LogFile\MACRO RUN\" & ProductID & "\" & nFile) For Input As #1  '�}�Ҥ�r��
'
'                    Line Input #1, str              '�v��Ū��
'                    b = " - "
'                    message = Split(str, b)
'                    For i = 0 To UBound(message)
'                        MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
'                        MDIForm1.Text1.Font.Size = 13
'                    Next i
'
'                    Close #1
'                End If
'            End If
'         End If
'        nFile = Dir
'    Loop
'--------------------------------D��
nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\Final Macro�Ϥ�\*.txt")  '�]�w�Q�n�B�z���ؿ��� C:\123, �B�z���ɮװ��ɦW�� *.txt

    Do While Len(nFile)
        If nFile <> "" Then
            File = Split(nFile, "_")
            aFile = Split(nFile, ".")
            If FSO.FileExists("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\Final Macro�Ϥ�\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".txt") Then
                If File(1) <> "" Then
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Text1.Text = ""             '�M�����e
                        
                        Open ("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\Final Macro�Ϥ�\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".txt") For Input As #1  '�}�Ҥ�r��
    
                        Line Input #1, str              '�v��Ū��
                        b = " - "
                        message = Split(str, b)
                        For i = 0 To UBound(message)
                            MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                            MDIForm1.Text1.Font.Size = 13
                        Next i
    
                        Close #1
                    End If
                End If
            End If
            
        End If
        nFile = Dir
    Loop
            

   MDIForm1.Label1 = Down & "/" & j & "  ��"

  
     
finish:
exA_GlassID = GlassID
End Sub
