Attribute VB_Name = "�Ϥ�"
Private bFirstIn As Boolean
Private nWidth, nHeight As Single



Sub Prdpicture(ByVal Down, ByVal g_sProductID As String)

    On Error Resume Next
    Dim nFile As String
    Dim str As String
    Dim b As String
    Dim message() As String
    Dim FSO As FileSystemObject

    'ProductID = "FGF640XXQ1"
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

    
    
    'MUIN
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\CV Macro\" & ProductID & "\*.jpg")
    
    'MUIN 20
    'nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\*.jpg")  '�]�w�Q�n�B�z���ؿ��� C:\123, �B�z���ɮװ��ɦW�� *.txt

a = a
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

'--------------------------------�{�ɥ[�j��
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\CV Macro\*.txt")
    Do While Len(nFile) '�p���Ƨ������X�i�Ϥ�
a = a
        If nFile <> "" Then
            aFile = Split(nFile, ".")
   
            'Macro�q��
            If aFile(0) <> "" Then
        
                j = j + 1
                If j = Down Then
                    MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\CV Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".jpg")
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
'--------------------------------�{�ɥ[�j��

    'MUIN
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\CV Macro\" & ProductID & "\*.jpg")
    
    'MUIN 20
    'nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\*.jpg")
    Do While Len(nFile) '�p���Ƨ������X�i�Ϥ�
a = a
        If nFile <> "" Then
            File = Split(nFile, "_")

            If (CoaterID = "13" And File(1) = "L") Or (CoaterID = "12" And File(1) = "R") Or File(1) = "Coater" Then
                j = j + 1
                If j = Down Then
                    'MUIN
                    'MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\CV Macro\" & ProductID & "\" & nFile)
                
                    'MUIN 20
                    MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\CV Macro\" & ProductID & "\" & nFile)
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
    j = 0
    a = a
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

'--------------------------------�{�ɥ[�j��
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\CV Macro\*.txt")
    Do While Len(nFile) '�p���Ƨ������X�i�Ϥ�
a = a
        If nFile <> "" Then
            aFile = Split(nFile, ".")
            
            '�q��
            If aFile(0) <> "" Then
        
                j = j + 1
                If j = Down Then
                    MDIForm1.Text1.Text = ""
                    Open ("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\CV Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".txt") For Input As #3
                    
                    Line Input #3, str              '�v��Ū��
                    b = " - "
                    message = Split(str, b)
                    For i = 0 To UBound(message)
                        MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                        MDIForm1.Text1.Font.Size = 13
                    Next i
    
                    Close #3
                End If
            End If
        End If
        
        nFile = Dir
    Loop
'--------------------------------�{�ɥ[�j��

'MUIN
nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\CV Macro\" & ProductID & "\*.txt")

'MUIN 20
'nFile = Dir("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\SAMP Macro\" & ProductID & "\*.txt")  '�]�w�Q�n�B�z���ؿ��� C:\123, �B�z���ɮװ��ɦW�� *.txt

    Do While Len(nFile)
        If nFile <> "" Then
            File = Split(nFile, "_")

            If (CoaterID = "13" And File(1) = "L") Or (CoaterID = "12" And File(1) = "R") Or File(1) = "Coater" Then
                j = j + 1
                If j = Down Then
                    MDIForm1.Text1.Text = ""             '�M�����e
                    
                    'MUIN
                    'Open ("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\CV Macro\" & ProductID & "\" & nFile) For Input As #1 '�}�Ҥ�r��
                    
                    'MUIN 20
                    Open ("\\10.91.40.40\fabsh$\cf5\�s�y��\�ժ��M��\�C�B��\�Ϥ�\CV Macro\" & ProductID & "\" & nFile) For Input As #1 '�}�Ҥ�r��

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
        nFile = Dir
    Loop
            
            
   a = a

   MDIForm1.Label1 = Down & "/" & j & "  ��"

   
    
     
finish:
exA_GlassID = GlassID
End Sub
