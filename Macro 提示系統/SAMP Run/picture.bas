Attribute VB_Name = "�Ϥ�"
Private bFirstIn As Boolean
Private nWidth, nHeight As Single



Sub Prdpicture(ByVal Down, ByVal g_sProductID As String)
    'On Error Resume Next
    On Error GoTo Err
    
    
    Dim nFile As String
    Dim str As String
    Dim b As String
    Dim message() As String
    Dim fso As FileSystemObject


    'ProductID = "OGH313BFS1"
    'OperationID = "5290"
    'Down = 2
    
    'Macro07
    'LineID = "07"

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

    If ProductID <> "" Then
        Open "I:\�s�y��\�ժ��M��\�C�B��\�Ϥ�\�L����\�T�{�s�u.txt" For Output As #3
        Print #3, "Fab�Ѧ��s�u"
        Close #3
    End If
    
    
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\*.jpg")  '�]�w�Q�n�B�z���ؿ��� C:\123, �B�z���ɮװ��ɦW�� *.txt

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
    nFile = Dir("D:\LogFile\MACRO RUN\" & ProductID & "\*.jpg")
        Do While Len(nFile)
    If nFile <> "" Then
            File = Split(nFile, "_")
            If File(1) <> "" Then   'Macro�q��
                j = j + 1
                If j = Down Then
                    MDIForm1.Image1.Picture = LoadPicture("D:\LogFile\MACRO RUN\" & ProductID & "\" & nFile)
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
    nFile = Dir
    Loop
    nFile = ""
'--------------------------------D��
'--------------------------------�{�ɥ[�j��
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\*.txt")
    Do While Len(nFile) '�p���Ƨ������X�i�Ϥ�

        If nFile <> "" Then
            aFile = Split(nFile, ".")
   
            'Macro�q��
            If aFile(0) <> "" Then
                Set fso = CreateObject("Scripting.FileSystemObject")
                If fso.FileExists("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".jpg") Then
                j = j + 1
                    If j = Down Then
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".jpg")
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
    nFile = ""
'--------------------------------�{�ɥ[�j��
'--------------------------------�{�ɥ[�jALL��i
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\ALL\*.jpg")
    Do While Len(nFile) '�p���Ƨ������X�i�Ϥ�

        If nFile <> "" Then
            j = j + 1
            If j = Down Then
                MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\ALL\" & nFile)
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
        nFile = Dir
    Loop
    nFile = ""

'--------------------------------�{�ɥ[�jALL��i



    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\*.jpg")
    Do While Len(nFile) '�p���Ƨ������X�i�Ϥ�

        If nFile <> "" Then
            File = Split(nFile, "_")
            'Macro07����
            'If File(1) = "07" Or File(1) = "���p�I.jpg" Or File(1) = "����.jpg" Or File(1) = Left(OperationID, 2) Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then
            
            'Macro�q��
            If File(1) <> "" Then
                If File(1) = Left(OperationID, 2) Then
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\" & nFile)
                        xFile = Split(nFile, ".")
                        MDIForm1.Label4 = xFile(0)
                        MDIForm1.Label2 = "ProductID:  " & ProductID
                        MDIForm1.Label3 = "Operation ID:  " & OperationID
                        MDIForm1.Label1.Font.Size = 15
                        MDIForm1.Label2.Font.Size = 15
                        MDIForm1.Label3.Font.Size = 15
                        MDIForm1.Label4.Font.Size = 15
                    End If

                ElseIf CoaterID = "12" And File(1) = "R" Then
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\" & nFile)
                        xFile = Split(nFile, ".")
                        MDIForm1.Label4 = xFile(0)
                        MDIForm1.Label2 = "ProductID:  " & ProductID
                        MDIForm1.Label3 = "Operation ID:  " & OperationID
                        MDIForm1.Label1.Font.Size = 15
                        MDIForm1.Label2.Font.Size = 15
                        MDIForm1.Label3.Font.Size = 15
                        MDIForm1.Label4.Font.Size = 15
                    End If

                ElseIf CoaterID = "13" And File(1) = "L" Then
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\" & nFile)
                        xFile = Split(nFile, ".")
                        MDIForm1.Label4 = xFile(0)
                        MDIForm1.Label2 = "ProductID:  " & ProductID
                        MDIForm1.Label3 = "Operation ID:  " & OperationID
                        MDIForm1.Label1.Font.Size = 15
                        MDIForm1.Label2.Font.Size = 15
                        MDIForm1.Label3.Font.Size = 15
                        MDIForm1.Label4.Font.Size = 15
                    End If

                ElseIf File(1) <> "R" Then
                    If File(1) <> "L" Then
                    j = j + 1
                        If j = Down Then
                            MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\" & nFile)
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
                
                'Macro07����
'                ElseIf File(1) = LineID Then
'                    j = j + 1
'                    If j = Down Then
'                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\" & nFile)
'                        xFile = Split(nFile, ".")
'                        MDIForm1.Label4 = xFile(0)
'                        MDIForm1.Label2 = "ProductID:  " & ProductID
'                        MDIForm1.Label3 = "Operation ID:  " & OperationID
'                        MDIForm1.Label1.Font.Size = 15
'                        MDIForm1.Label2.Font.Size = 15
'                        MDIForm1.Label3.Font.Size = 15
'                        MDIForm1.Label4.Font.Size = 15
'                    End If
                
                End If
            End If
    
        End If
        
        nFile = Dir
    Loop
    j = 0

    
    
    
    
    
    
    
    
    
nFile = Dir("D:\LogFile\MACRO RUN\" & ProductID & "\*.txt")
'--------------------------------D��
     Do While Len(nFile)

        If nFile <> "" Then
            File = Split(nFile, "_")
            If File(1) <> "" Then 'Macro�q��

                j = j + 1
                If j = Down Then
                    MDIForm1.Text1.Text = ""             '�M�����e
                    Open ("D:\LogFile\MACRO RUN\" & ProductID & "\" & nFile) For Input As #1  '�}�Ҥ�r��
    
                    Line Input #1, str              '�v��Ū��
                    b = "_"
                    message = Split(str, b)
                    For i = 0 To UBound(message)
                        MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                        MDIForm1.Text1.Font.Size = 13
                    Next i
    
                    Close #1
                End If
            End If
         End If
        nFile = Dir
    Loop
'--------------------------------D��
'--------------------------------�{�ɥ[�j��
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\*.txt")
    Do While Len(nFile) '�p���Ƨ������X�i�Ϥ�
a = a
        If nFile <> "" Then
            aFile = Split(nFile, ".")
            
            'Macro�q��
            If aFile(0) <> "" Then
                Set fso = CreateObject("Scripting.FileSystemObject")
                If fso.FileExists("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".txt") Then
        
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Text1.Text = ""
                        Open ("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".txt") For Input As #3
                        
                        Line Input #3, str              '�v��Ū��
                        b = "_"
                        message = Split(str, b)
                        For i = 0 To UBound(message)
                            MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                            MDIForm1.Text1.Font.Size = 13
                        Next i
        
                        Close #3
                    End If
                End If
            End If
        End If
        
        nFile = Dir
    Loop
'--------------------------------�{�ɥ[�j��
'--------------------------------�{�ɥ[�jALL��i

    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\ALL\*.txt")
    Do While Len(nFile) '�p���Ƨ������X�i�Ϥ�

        If nFile <> "" Then
            j = j + 1
            If j = Down Then
                MDIForm1.Text1.Text = ""
                Open ("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\ALL\" & nFile) For Input As #4
                Line Input #4, str              '�v��Ū��
                b = "_"
                message = Split(str, b)
                For i = 0 To UBound(message)
                    MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                    MDIForm1.Text1.Font.Size = 13
                Next i
                Close #4
            End If
        End If
        nFile = Dir
    Loop
    nFile = ""

'--------------------------------�{�ɥ[�jALL��i
nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\*.txt")  '�]�w�Q�n�B�z���ؿ��� C:\123, �B�z���ɮװ��ɦW�� *.txt
       
    Do While Len(nFile)
        If nFile <> "" Then
            File = Split(nFile, "_")
           
            'Macro�q��
            If File(1) <> "" Then
                
                If File(1) = Left(OperationID, 2) Then
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Text1.Text = ""             '�M�����e
                        Open ("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\" & nFile) For Input As #2 '�}�Ҥ�r��
        
                        Line Input #2, str              '�v��Ū��
                        b = "_"
                        message = Split(str, b)
                        For i = 0 To UBound(message)
                            MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                            MDIForm1.Text1.Font.Size = 13
                        Next i
        
                        Close #2
                    End If
                
                
                ElseIf CoaterID = "12" And File(1) = "R" Then
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Text1.Text = ""             '�M�����e
                        Open ("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\" & nFile) For Input As #2 '�}�Ҥ�r��
        
                        Line Input #2, str              '�v��Ū��
                        b = "_"
                        message = Split(str, b)
                        For i = 0 To UBound(message)
                            MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                            MDIForm1.Text1.Font.Size = 13
                        Next i
        
                        Close #2
                    End If
                
                
                ElseIf CoaterID = "13" And File(1) = "L" Then
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Text1.Text = ""             '�M�����e
                        Open ("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\" & nFile) For Input As #2 '�}�Ҥ�r��
        
                        Line Input #2, str              '�v��Ū��
                        b = "_"
                        message = Split(str, b)
                        For i = 0 To UBound(message)
                            MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                            MDIForm1.Text1.Font.Size = 13
                        Next i
        
                        Close #2
                    End If
                
                
                ElseIf File(1) <> "R" Then
                    If File(1) <> "L" Then
                    j = j + 1
                        If j = Down Then
                            MDIForm1.Text1.Text = ""             '�M�����e
                            Open ("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\" & nFile) For Input As #2 '�}�Ҥ�r��
            
                            Line Input #2, str              '�v��Ū��
                            b = "_"
                            message = Split(str, b)
                            For i = 0 To UBound(message)
                                MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                                MDIForm1.Text1.Font.Size = 13
                            Next i
            
                            Close #2
                        End If
                    End If
                    
                'Macro07����
'                ElseIf File(1) = LineID Then
'                    j = j + 1
'                    If j = Down Then
'                        MDIForm1.Text1.Text = ""             '�M�����e
'                        Open ("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\" & nFile) For Input As #2 '�}�Ҥ�r��
'
'                        Line Input #2, str              '�v��Ū��
'                        b = "_"
'                        message = Split(str, b)
'                        For i = 0 To UBound(message)
'                            MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
'                            MDIForm1.Text1.Font.Size = 13
'                        Next i
'
'                        Close #2
'                    End If
'
'                End If
'            End If
            


        End If
        nFile = Dir
    Loop
            
            


   
   MDIForm1.Label1 = Down & "/" & j & "  ��"
   'exProduct = ProductID
   
   GoTo finish
     


Err:
MsgBox "Fab�ѳs�u���ѡA�Э��s�s�u"

Exit Sub
finish:
exA_GlassID = GlassID


End Sub
