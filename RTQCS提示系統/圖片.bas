Attribute VB_Name = "圖片"
Sub Prodpicture()
 Dim str As String
 Dim fso As FileSystemObject
On Error GoTo ErrorHandle
'    If ex_Product_1 = Product_1 And ex_Product_1 <> "" Then
'        GoTo stage2
'    End If

    j = 0
    
    If Oper_1 = "" Then
        nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\宣導\*.jpg")
        Do While Len(nFile)
            tFile = Split(nFile, ".")
            j = j + 1
            If (j * 15) = change Then
                MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\宣導\" & nFile)
                'txt
                Set fso = CreateObject("Scripting.FileSystemObject")
                If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\宣導\" & tFile(0) & ".txt") Then
                    MDIForm1.Text1 = ""
                    Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\宣導\" & tFile(0) & ".txt") For Input As #3
                    Line Input #3, str
                    message = Split(str, "_")
                    For i = 0 To UBound(message)
                        MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                        MDIForm1.Text1.Font.Size = 13
                    Next i
                    Close #3
                Else
                    MDIForm1.Text1 = ""
                End If
                
            End If
        nFile = Dir
        Loop
       a = a

        
        
        If nFile = "" And change > (j * 15) Then
            change = 1
        End If
    End If
    
If Oper_1 <> "" Then
    If Choose_which1 = 1 Then
        nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_1 & "\" & Left(Operation_1, 2) & "00\*.jpg")
        If ex_Operation_1 <> Operation_1 Then
            MDIForm1.Image1.Picture = LoadPicture("")
            MDIForm1.Text1 = ""
            change1_1xxx = 9
            change1_2xxx = 9
            change1_5xxx = 9
            change1_7xxx = 9
            change1_8xxx = 9
            change1_DMRV = 9
        ElseIf nFile <> "" Then
            If Left(Operation_1, 2) = 72 Or Left(Operation_1, 2) = 73 Or Left(Operation_1, 2) = 23 Or Left(Operation_1, 2) = 33 Then
                MDIForm1.Image1.Picture = LoadPicture("")
                MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_1 & "\" & Left(Operation_1, 2) & "00\" & nFile)
            End If
        End If
    
    
    Operation1_1xxx = 0
    Operation1_2xxx = 0
    Operation1_5xxx = 0
    Operation1_7xxx = 0
    Operation1_8xxx = 0
'------------------

        '到站點跑該站宣導
        If Left(Operation_1, 1) = 1 Then
            nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\BM宣導\BM*.jpg")
            change1_1xxx = change1_1xxx + 1
            Do While Len(nFile)
                tFile = Split(nFile, ".")
                Operation1_1xxx = Operation1_1xxx + 1
                If (Operation1_1xxx * 10) = change1_1xxx Then
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\BM宣導\" & nFile) Then
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\BM宣導\" & nFile)
                        Set fso = Nothing
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\BM宣導\" & tFile(0) & ".txt") Then
                            MDIForm1.Text1 = ""
                            Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\BM宣導\" & tFile(0) & ".txt") For Input As #1
                            Line Input #1, str              '逐行讀取
                            message = Split(str, "_")
                            For i = 0 To UBound(message)
                                MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                                MDIForm1.Text1.Font.Size = 13
                            Next i
                        Close #1
                        Else
                            MDIForm1.Text1 = ""
                        End If
                    End If
                End If
                nFile = Dir
            Loop
            
        ElseIf Left(Operation_1, 1) = 2 Or Left(Operation_1, 1) = 3 Or Left(Operation_1, 1) = 4 Then
            If Left(Operation_1, 2) = 23 Or Left(Operation_1, 2) = 33 Then
                YFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_1 & "\" & Left(Operation_1, 2) & "00\*.jpg")
                If YFile = "" Then
                    nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\RGB*.jpg")
                    change1_2xxx = change1_2xxx + 1
                    Do While Len(nFile)
                        tFile = Split(nFile, ".")
                        Operation1_2xxx = Operation1_2xxx + 1
                        If (Operation1_2xxx * 10) = change1_2xxx Then
                            Set fso = CreateObject("Scripting.FileSystemObject")
                            If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & nFile) Then
                                MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & nFile)
                                Set fso = Nothing
                                Set fso = CreateObject("Scripting.FileSystemObject")
                                If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & tFile(0) & ".txt") Then
                                    MDIForm1.Text1 = ""
                                    Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & tFile(0) & ".txt") For Input As #1
                                    Line Input #1, str              '逐行讀取
                                    message = Split(str, "_")
                                    For i = 0 To UBound(message)
                                        MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                                        MDIForm1.Text1.Font.Size = 13
                                    Next i
                                Close #1
                                Else
                                    MDIForm1.Text1 = ""
                                End If
                            End If
                        End If
                        nFile = Dir
                    Loop
                End If
            End If
            
            If Left(Operation_1, 2) <> 23 Then
                If Left(Operation_1, 2) <> 33 Then
                    nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\RGB*.jpg")
                    change1_2xxx = change1_2xxx + 1
                    Do While Len(nFile)
                        tFile = Split(nFile, ".")
                        Operation1_2xxx = Operation1_2xxx + 1
                        If (Operation1_2xxx * 10) = change1_2xxx Then
                            Set fso = CreateObject("Scripting.FileSystemObject")
                            If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & nFile) Then
                                MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & nFile)
                                Set fso = Nothing
                                Set fso = CreateObject("Scripting.FileSystemObject")
                                If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & tFile(0) & ".txt") Then
                                    MDIForm1.Text1 = ""
                                    Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & tFile(0) & ".txt") For Input As #1
                                    Line Input #1, str              '逐行讀取
                                    message = Split(str, "_")
                                    For i = 0 To UBound(message)
                                        MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                                        MDIForm1.Text1.Font.Size = 13
                                    Next i
                                Close #1
                                Else
                                    MDIForm1.Text1 = ""
                                End If
                            End If
                        End If
                        nFile = Dir
                    Loop
                End If
            End If
                
                
        ElseIf Left(Operation_1, 1) = 5 Then
            nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\ITO宣導\ITO*.jpg")
            change1_5xxx = change1_5xxx + 1
            Do While Len(nFile)
                tFile = Split(nFile, ".")
                Operation1_5xxx = Operation1_5xxx + 1
                If (Operation1_5xxx * 10) = change1_5xxx Then
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\ITO宣導\" & nFile) Then
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\ITO宣導\" & nFile)
                        Set fso = Nothing
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\ITO宣導\" & tFile(0) & ".txt") Then
                            MDIForm1.Text1 = ""
                            Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\ITO宣導\" & tFile(0) & ".txt") For Input As #1
                            Line Input #1, str              '逐行讀取
                            message = Split(str, "_")
                            For i = 0 To UBound(message)
                                MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                                MDIForm1.Text1.Font.Size = 13
                            Next i
                        Close #1
                        Else
                            MDIForm1.Text1 = ""
                        End If
                    End If
                End If
                nFile = Dir
            Loop
            
        ElseIf Left(Operation_1, 1) = 6 Or Left(Operation_1, 1) = 9 Then
        
        ElseIf Left(Operation_1, 1) = 7 Then
            If Left(Operation_1, 2) = 72 Or Left(Operation_1, 2) = 73 Then
                YFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_1 & "\" & Left(Operation_1, 2) & "00\*.jpg")
                If YFile = "" Then
                    nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\PS*.jpg")
                    change1_7xxx = change1_7xxx + 1
                    Do While Len(nFile)
                        tFile = Split(nFile, ".")
                        Operation1_7xxx = Operation1_7xxx + 1
                        If (Operation1_7xxx * 10) = change1_7xxx Then
                            Set fso = Nothing
                            Set fso = CreateObject("Scripting.FileSystemObject")
                            If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & nFile) Then
                                MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & nFile)
                                Set fso = Nothing
                                Set fso = CreateObject("Scripting.FileSystemObject")
                                If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & tFile(0) & ".txt") Then
                                    MDIForm1.Text1 = ""
                                    Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & tFile(0) & ".txt") For Input As #1
                                    Line Input #1, str              '逐行讀取
                                    message = Split(str, "_")
                                    For i = 0 To UBound(message)
                                        MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                                        MDIForm1.Text1.Font.Size = 13
                                    Next i
                                Close #1
                                Else
                                    MDIForm1.Text1 = ""
                                End If
                            End If
                        End If
                        nFile = Dir
                    Loop
                End If
                
            End If
            
            If Left(Operation_1, 2) <> 72 Then
                If Left(Operation_1, 2) <> 73 Then
                    nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\PS*.jpg")
                    change1_7xxx = change1_7xxx + 1
                    Do While Len(nFile)
                        tFile = Split(nFile, ".")
                        Operation1_7xxx = Operation1_7xxx + 1
                        If (Operation1_7xxx * 10) = change1_7xxx Then
                            Set fso = CreateObject("Scripting.FileSystemObject")
                            If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & nFile) Then
                                MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & nFile)
                                Set fso = Nothing
                                Set fso = CreateObject("Scripting.FileSystemObject")
                                If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & tFile(0) & ".txt") Then
                                    MDIForm1.Text1 = ""
                                    Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & tFile(0) & ".txt") For Input As #1
                                    Line Input #1, str              '逐行讀取
                                    message = Split(str, "_")
                                    For i = 0 To UBound(message)
                                        MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                                        MDIForm1.Text1.Font.Size = 13
                                    Next i
                                Close #1
                                Else
                                    MDIForm1.Text1 = ""
                                End If
                            End If
                        End If
                        nFile = Dir
                    Loop
                End If
            End If
            
        ElseIf Left(Operation_1, 1) = 8 Then
            nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\OC宣導\OC*.jpg")
            change1_8xxx = change1_8xxx + 1
            Do While Len(nFile)
                tFile = Split(nFile, ".")
                Operation1_8xxx = Operation1_8xxx + 1
                If (Operation1_8xxx * 10) = change1_8xxx Then
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\OC宣導\" & nFile) Then
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\OC宣導\" & nFile)
                        Set fso = Nothing
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\OC宣導\" & tFile(0) & ".txt") Then
                            MDIForm1.Text1 = ""
                            Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\OC宣導\" & tFile(0) & ".txt") For Input As #1
                            Line Input #1, str              '逐行讀取
                            message = Split(str, "_")
                            For i = 0 To UBound(message)
                                MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                                MDIForm1.Text1.Font.Size = 13
                            Next i
                        Close #1
                        Else
                            MDIForm1.Text1 = ""
                        End If
                    End If
                End If
                nFile = Dir
            Loop
        End If
        
    ElseIf Choose_which1 = 2 Then
    
    Operation1_DMRV = 0

        nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\DMRV宣導\DM*.jpg")
        change1_DMRV = change1_DMRV + 1
            Do While Len(nFile)
                tFile = Split(nFile, ".")
                Operation1_DMRV = Operation1_DMRV + 1
                If (Operation1_DMRV * 10) = change1_DMRV Then
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\DMRV宣導\" & nFile) Then
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\DMRV宣導\" & nFile)
                        Set fso = Nothing
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\DMRV宣導\" & tFile(0) & ".txt") Then
                            MDIForm1.Text1 = ""
                            Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\DMRV宣導\" & tFile(0) & ".txt") For Input As #1
                            Line Input #1, str              '逐行讀取
                            message = Split(str, "_")
                            For i = 0 To UBound(message)
                                MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                                MDIForm1.Text1.Font.Size = 13
                            Next i
                        Close #1
                        Else
                            MDIForm1.Text1 = ""
                        End If
                    End If
                End If
                nFile = Dir
            Loop
            
    End If


        
'--------------------------------------------------------------------------
        'txt
        aFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_1 & "\" & Left(Operation_1, 2) & "00\*.txt")
        If ex_Operation_1 <> Operation_1 Then
            'MDIForm1.Text1 = ""
        ElseIf aFile <> "" Then
            If Left(Operation_1, 2) = 72 Or Left(Operation_1, 2) = 73 Or Left(Operation_1, 2) = 23 Or Left(Operation_1, 2) = 33 Then
                MDIForm1.Text1 = ""
                Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_1 & "\" & Left(Operation_1, 2) & "00\" & aFile) For Input As #1
                Line Input #1, str              '逐行讀取
                message = Split(str, "_")
                For i = 0 To UBound(message)
                    MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                    MDIForm1.Text1.Font.Size = 13
                Next i
        
                Close #1
            End If
        End If
        
End If
    a = a
        If Operation1_1xxx <> 0 And change1_1xxx > (Operation1_1xxx * 10) Then
            change1_1xxx = 1
        End If
        If Operation1_2xxx <> 0 And change1_2xxx > (Operation1_2xxx * 10) Then
            change1_2xxx = 1
        End If
        If Operation1_5xxx <> 0 And change1_5xxx > (Operation1_5xxx * 10) Then
            change1_5xxx = 1
        End If
        If Operation1_7xxx <> 0 And change1_7xxx > (Operation1_7xxx * 10) Then
            change1_7xxx = 1
        End If
        If Operation1_8xxx <> 0 And change1_8xxx > (Operation1_8xxx * 10) Then
            change1_8xxx = 1
        End If
        If Operation1_DMRV <> 0 And change1_DMRV > (Operation1_DMRV * 10) Then
            change1_DMRV = 1
        End If
'---------------------------------------------------
stage2:



    Operation2_1xxx = 0
    Operation2_2xxx = 0
    Operation2_5xxx = 0
    Operation2_7xxx = 0
    Operation2_8xxx = 0
    
'    If ex_Product_2 = Product_2 And ex_Product_2 <> "" Then
'        GoTo finish
'    End If
    
    j = 0
    If Oper_2 = "" Then
        nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\宣導\*.jpg")
        Do While Len(nFile)
        tFile = Split(nFile, ".")
            j = j + 1
            If (j * 15) = change Then
                MDIForm1.Image2.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\宣導\" & nFile)
                'txt
                Set fso = CreateObject("Scripting.FileSystemObject")
                If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\宣導\" & tFile(0) & ".txt") Then
                    MDIForm1.Text2 = ""
                    Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\宣導\" & tFile(0) & ".txt") For Input As #4
                    Line Input #4, str
                    message = Split(str, "_")
                    For i = 0 To UBound(message)
                        MDIForm1.Text2.Text = MDIForm1.Text2.Text & message(i) & vbCrLf
                        MDIForm1.Text2.Font.Size = 13
                    Next i
                    Close #4
                Else
                    MDIForm1.Text2 = ""
                End If
                
            End If
        nFile = Dir
        Loop
        If nFile = "" And change > (j * 15) Then
            change = 1
        End If
    End If

    
If Oper_2 <> "" Then
    If Choose_which2 = 1 Then
    
        nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_2 & "\" & Left(Operation_2, 2) & "00\*.jpg")
        If ex_Operation_2 <> Operation_2 Then
            MDIForm1.Image2.Picture = LoadPicture("")
            MDIForm1.Text2 = ""
            change2_1xxx = 9
            change2_2xxx = 9
            change2_5xxx = 9
            change2_7xxx = 9
            change2_8xxx = 9
            change2_DMRV = 9
        ElseIf nFile <> "" Then
            If Left(Operation_2, 2) = 72 Or Left(Operation_2, 2) = 73 Or Left(Operation_2, 2) = 23 Or Left(Operation_2, 2) = 33 Then
                MDIForm1.Image2.Picture = LoadPicture("")
                MDIForm1.Image2.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_2 & "\" & Left(Operation_2, 2) & "00\" & nFile)
            End If
        End If
'------------------
        '到站點跑該站宣導
        If Left(Operation_2, 1) = 1 Then
        a = a
            nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\BM宣導\BM*.jpg")
            change2_1xxx = change2_1xxx + 1
            Do While Len(nFile)
                tFile = Split(nFile, ".")
                Operation2_1xxx = Operation2_1xxx + 1
                If (Operation2_1xxx * 10) = change2_1xxx Then
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\BM宣導\" & nFile) Then
                        MDIForm1.Image2.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\BM宣導\" & nFile)
                        Set fso = Nothing
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\BM宣導\" & tFile(0) & ".txt") Then
                            MDIForm1.Text2 = ""
                            Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\BM宣導\" & tFile(0) & ".txt") For Input As #2
                            Line Input #2, str              '逐行讀取
                            message = Split(str, "_")
                            For i = 0 To UBound(message)
                                MDIForm1.Text2.Text = MDIForm1.Text2.Text & message(i) & vbCrLf
                                MDIForm1.Text2.Font.Size = 13
                            Next i
                        Close #2
                        Else
                            MDIForm1.Text2 = ""
                        End If
                    End If
                End If
                nFile = Dir
            Loop
                    
                    
                    
                    
        ElseIf Left(Operation_2, 1) = 2 Or Left(Operation_2, 1) = 3 Or Left(Operation_2, 1) = 4 Then
            If Left(Operation_2, 2) = 23 Or Left(Operation_2, 2) = 33 Then
                YFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_2 & "\" & Left(Operation_2, 2) & "00\*.jpg")
                If YFile = "" Then
                    nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\RGB*.jpg")
                    change2_2xxx = change2_2xxx + 1
                    Do While Len(nFile)
                        tFile = Split(nFile, ".")
                        Operation2_2xxx = Operation2_2xxx + 1
                        If (Operation2_2xxx * 10) = change2_2xxx Then
                            Set fso = CreateObject("Scripting.FileSystemObject")
                            If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & nFile) Then
                                MDIForm1.Image2.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & nFile)
                                Set fso = Nothing
                                Set fso = CreateObject("Scripting.FileSystemObject")
                                If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & tFile(0) & ".txt") Then
                                    MDIForm1.Text2 = ""
                                    Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & tFile(0) & ".txt") For Input As #2
                                    Line Input #2, str              '逐行讀取
                                    message = Split(str, "_")
                                    For i = 0 To UBound(message)
                                        MDIForm1.Text2.Text = MDIForm1.Text2.Text & message(i) & vbCrLf
                                        MDIForm1.Text2.Font.Size = 13
                                    Next i
                                Close #2
                                Else
                                    MDIForm1.Text2 = ""
                                End If
                            End If
                        End If
                        nFile = Dir
                    Loop
                End If
            End If
            
            If Left(Operation_2, 2) <> 23 Then
                If Left(Operation_2, 2) <> 33 Then
                    nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\RGB*.jpg")
                    change2_2xxx = change2_2xxx + 1
                    Do While Len(nFile)
                        tFile = Split(nFile, ".")
                        Operation2_2xxx = Operation2_2xxx + 1
                        If (Operation2_2xxx * 10) = change2_2xxx Then
                            Set fso = CreateObject("Scripting.FileSystemObject")
                            If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & nFile) Then
                                MDIForm1.Image2.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & nFile)
                                Set fso = Nothing
                                Set fso = CreateObject("Scripting.FileSystemObject")
                                If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & tFile(0) & ".txt") Then
                                    MDIForm1.Text2 = ""
                                    Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\RGB宣導\" & tFile(0) & ".txt") For Input As #2
                                    Line Input #2, str              '逐行讀取
                                    message = Split(str, "_")
                                    For i = 0 To UBound(message)
                                        MDIForm1.Text2.Text = MDIForm1.Text2.Text & message(i) & vbCrLf
                                        MDIForm1.Text2.Font.Size = 13
                                    Next i
                                Close #2
                                Else
                                    MDIForm1.Text2 = ""
                                End If
                            End If
                        End If
                        nFile = Dir
                    Loop
                End If
            End If
            
            
        ElseIf Left(Operation_2, 1) = 5 Then
            nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\ITO宣導\ITO*.jpg")
            change2_5xxx = change2_5xxx + 1
            Do While Len(nFile)
                tFile = Split(nFile, ".")
                Operation2_5xxx = Operation2_5xxx + 1
                If (Operation2_5xxx * 10) = change2_5xxx Then
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\ITO宣導\" & nFile) Then
                        MDIForm1.Image2.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\ITO宣導\" & nFile)
                        Set fso = Nothing
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\ITO宣導\" & tFile(0) & ".txt") Then
                            MDIForm1.Text2 = ""
                            Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\ITO宣導\" & tFile(0) & ".txt") For Input As #2
                            Line Input #2, str              '逐行讀取
                            message = Split(str, "_")
                            For i = 0 To UBound(message)
                                MDIForm1.Text2.Text = MDIForm1.Text2.Text & message(i) & vbCrLf
                                MDIForm1.Text2.Font.Size = 13
                            Next i
                        Close #2
                        Else
                            MDIForm1.Text2 = ""
                        End If
                    End If
                End If
                nFile = Dir
            Loop
            
            
        ElseIf Left(Operation_2, 1) = 6 Or Left(Operation_2, 1) = 9 Then
        
        ElseIf Left(Operation_2, 1) = 7 Then
            If Left(Operation_2, 2) = 72 Or Left(Operation_2, 2) = 73 Then
                YFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_2 & "\" & Left(Operation_2, 2) & "00\*.jpg")
                If YFile = "" Then
                    nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\PS*.jpg")
                    change2_7xxx = change2_7xxx + 1
                    Do While Len(nFile)
                        tFile = Split(nFile, ".")
                        Operation2_7xxx = Operation2_7xxx + 1
                        If (Operation2_7xxx * 10) = change2_7xxx Then
                            Set fso = Nothing
                            Set fso = CreateObject("Scripting.FileSystemObject")
                            If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & nFile) Then
                                MDIForm1.Image2.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & nFile)
                                Set fso = Nothing
                                Set fso = CreateObject("Scripting.FileSystemObject")
                                If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & tFile(0) & ".txt") Then
                                    MDIForm1.Text2 = ""
                                    Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & tFile(0) & ".txt") For Input As #2
                                    Line Input #2, str              '逐行讀取
                                    message = Split(str, "_")
                                    For i = 0 To UBound(message)
                                        MDIForm1.Text2.Text = MDIForm1.Text2.Text & message(i) & vbCrLf
                                        MDIForm1.Text2.Font.Size = 13
                                    Next i
                                Close #2
                                Else
                                    MDIForm1.Text2 = ""
                                End If
                            End If
                        End If
                        nFile = Dir
                    Loop
                End If
                
            End If
            
            If Left(Operation_2, 2) <> 72 Then
                If Left(Operation_2, 2) <> 73 Then
                    nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\PS*.jpg")
                    change2_7xxx = change2_7xxx + 1
                    Do While Len(nFile)
                        tFile = Split(nFile, ".")
                        Operation2_7xxx = Operation2_7xxx + 1
                        If (Operation2_7xxx * 10) = change2_7xxx Then
                            Set fso = CreateObject("Scripting.FileSystemObject")
                            If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & nFile) Then
                                MDIForm1.Image2.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & nFile)
                                Set fso = Nothing
                                Set fso = CreateObject("Scripting.FileSystemObject")
                                If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & tFile(0) & ".txt") Then
                                    MDIForm1.Text2 = ""
                                    Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\PS宣導\" & tFile(0) & ".txt") For Input As #2
                                    Line Input #2, str              '逐行讀取
                                    message = Split(str, "_")
                                    For i = 0 To UBound(message)
                                        MDIForm1.Text2.Text = MDIForm1.Text2.Text & message(i) & vbCrLf
                                        MDIForm1.Text2.Font.Size = 13
                                    Next i
                                Close #2
                                Else
                                    MDIForm1.Text2 = ""
                                End If
                            End If
                        End If
                        nFile = Dir
                    Loop
                End If
            End If
            
        ElseIf Left(Operation_2, 1) = 8 Then
            nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\OC宣導\OC*.jpg")
            change2_8xxx = change2_8xxx + 1
            Do While Len(nFile)
                tFile = Split(nFile, ".")
                Operation2_8xxx = Operation2_8xxx + 1
                If (Operation2_8xxx * 10) = change2_8xxx Then
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\OC宣導\" & nFile) Then
                        MDIForm1.Image2.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\OC宣導\" & nFile)
                        Set fso = Nothing
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\OC宣導\" & tFile(0) & ".txt") Then
                            MDIForm1.Text2 = ""
                            Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\OC宣導\" & tFile(0) & ".txt") For Input As #2
                            Line Input #2, str              '逐行讀取
                            message = Split(str, "_")
                            For i = 0 To UBound(message)
                                MDIForm1.Text2.Text = MDIForm1.Text2.Text & message(i) & vbCrLf
                                MDIForm1.Text2.Font.Size = 13
                            Next i
                        Close #2
                        Else
                            MDIForm1.Text2 = ""
                        End If
                    End If
                End If
                nFile = Dir
            Loop
              
        End If
        
        
    ElseIf Choose_which2 = 2 Then
    
        Operation2_DMRV = 0

        nFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\DMRV宣導\DM*.jpg")
        change2_DMRV = change2_DMRV + 1
            Do While Len(nFile)
                tFile = Split(nFile, ".")
                Operation2_DMRV = Operation2_DMRV + 1
                If (Operation2_DMRV * 10) = change2_DMRV Then
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\DMRV宣導\" & nFile) Then
                        MDIForm1.Image2.Picture = LoadPicture("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\DMRV宣導\" & nFile)
                        Set fso = Nothing
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        If fso.FileExists("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\DMRV宣導\" & tFile(0) & ".txt") Then
                            MDIForm1.Text2 = ""
                            Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\DMRV宣導\" & tFile(0) & ".txt") For Input As #2
                            Line Input #2, str              '逐行讀取
                            message = Split(str, "_")
                            For i = 0 To UBound(message)
                                MDIForm1.Text2.Text = MDIForm1.Text2.Text & message(i) & vbCrLf
                                MDIForm1.Text2.Font.Size = 13
                            Next i
                        Close #2
                        Else
                            MDIForm1.Text2 = ""
                        End If
                    End If
                End If
                nFile = Dir
            Loop
            
    
    
    End If
        
        
        
'--------------------------------------------------------------------------
        'txt
        aFile = Dir("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_2 & "\" & Left(Operation_2, 2) & "00\*.txt")
        If ex_Operation_2 <> Operation_2 Then
            'MDIForm1.Text2 = ""
        ElseIf aFile <> "" Then
            If Left(Operation_2, 2) = 72 Or Left(Operation_2, 2) = 73 Or Left(Operation_2, 2) = 23 Or Left(Operation_2, 2) = 33 Then
                MDIForm1.Text2 = ""
                Open ("\\10.91.40.40\fabsh$\CF5\製造部\組長專用\顏伯丞\小太陽提示\" & Product_2 & "\" & Left(Operation_2, 2) & "00\" & aFile) For Input As #2
                Line Input #2, str              '逐行讀取
                message = Split(str, "_")
                For i = 0 To UBound(message)
                    MDIForm1.Text2.Text = MDIForm1.Text2.Text & message(i) & vbCrLf
                    MDIForm1.Text2.Font.Size = 13
                Next i
        
                Close #2
            End If
        
        End If
        

    
        If Operation2_1xxx <> 0 And change2_1xxx > (Operation2_1xxx * 10) Then
            change2_1xxx = 1
        End If
        If Operation2_2xxx <> 0 And change2_2xxx > (Operation2_2xxx * 10) Then
            change2_2xxx = 1
        End If
        If Operation2_5xxx <> 0 And change2_5xxx > (Operation2_5xxx * 10) Then
            change2_5xxx = 1
        End If
        If Operation2_7xxx <> 0 And change2_7xxx > (Operation2_7xxx * 10) Then
            change2_7xxx = 1
        End If
        If Operation2_8xxx <> 0 And change2_8xxx > (Operation2_8xxx * 10) Then
            change2_8xxx = 1
        End If
        If Operation2_DMRV <> 0 And change2_DMRV > (Operation2_DMRV * 10) Then
            change2_DMRV = 1
        End If
End If
    
finish:

ex_Product_1 = Product_1
ex_Product_2 = Product_2
ex_Operation_1 = Operation_1
ex_Operation_2 = Operation_2

Exit Sub
ErrorHandle:
MsgBox "圖片錯誤"
End Sub
