Attribute VB_Name = "圖片"
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
        Open "I:\製造部\組長專用\顏伯丞\圖片\無提示\確認連線.txt" For Output As #3
        Print #3, "Fab槽有連線"
        Close #3
    End If
    
    
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\*.jpg")  '設定想要處理的目錄為 C:\123, 處理的檔案副檔名為 *.txt

    If nFile = "" Then  '要是沒有檔案 傳無提示
        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\無提示\無提示.jpg")
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
'--------------------------------D槽
    nFile = Dir("D:\LogFile\MACRO RUN\" & ProductID & "\*.jpg")
        Do While Len(nFile)
    If nFile <> "" Then
            File = Split(nFile, "_")
            If File(1) <> "" Then   'Macro通用
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
'--------------------------------D槽
'--------------------------------臨時加強檢
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\*.txt")
    Do While Len(nFile) '計算資料夾內有幾張圖片

        If nFile <> "" Then
            aFile = Split(nFile, ".")
   
            'Macro通用
            If aFile(0) <> "" Then
                Set fso = CreateObject("Scripting.FileSystemObject")
                If fso.FileExists("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".jpg") Then
                j = j + 1
                    If j = Down Then
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".jpg")
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
'--------------------------------臨時加強檢
'--------------------------------臨時加強ALL單張
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\ALL\*.jpg")
    Do While Len(nFile) '計算資料夾內有幾張圖片

        If nFile <> "" Then
            j = j + 1
            If j = Down Then
                MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\ALL\" & nFile)
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

'--------------------------------臨時加強ALL單張



    nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\*.jpg")
    Do While Len(nFile) '計算資料夾內有幾張圖片

        If nFile <> "" Then
            File = Split(nFile, "_")
            'Macro07條件
            'If File(1) = "07" Or File(1) = "六小點.jpg" Or File(1) = "版邊.jpg" Or File(1) = Left(OperationID, 2) Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then
            
            'Macro通用
            If File(1) <> "" Then
                If File(1) = Left(OperationID, 2) Then
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\" & nFile)
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
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\" & nFile)
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
                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\" & nFile)
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
                            MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\" & nFile)
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
                
                'Macro07條件
'                ElseIf File(1) = LineID Then
'                    j = j + 1
'                    If j = Down Then
'                        MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\" & nFile)
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
'--------------------------------D槽
     Do While Len(nFile)

        If nFile <> "" Then
            File = Split(nFile, "_")
            If File(1) <> "" Then 'Macro通用

                j = j + 1
                If j = Down Then
                    MDIForm1.Text1.Text = ""             '清除內容
                    Open ("D:\LogFile\MACRO RUN\" & ProductID & "\" & nFile) For Input As #1  '開啟文字檔
    
                    Line Input #1, str              '逐行讀取
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
'--------------------------------D槽
'--------------------------------臨時加強檢
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\*.txt")
    Do While Len(nFile) '計算資料夾內有幾張圖片
a = a
        If nFile <> "" Then
            aFile = Split(nFile, ".")
            
            'Macro通用
            If aFile(0) <> "" Then
                Set fso = CreateObject("Scripting.FileSystemObject")
                If fso.FileExists("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".txt") Then
        
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Text1.Text = ""
                        Open ("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".txt") For Input As #3
                        
                        Line Input #3, str              '逐行讀取
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
'--------------------------------臨時加強檢
'--------------------------------臨時加強ALL單張

    nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\ALL\*.txt")
    Do While Len(nFile) '計算資料夾內有幾張圖片

        If nFile <> "" Then
            j = j + 1
            If j = Down Then
                MDIForm1.Text1.Text = ""
                Open ("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\ALL\" & nFile) For Input As #4
                Line Input #4, str              '逐行讀取
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

'--------------------------------臨時加強ALL單張
nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\*.txt")  '設定想要處理的目錄為 C:\123, 處理的檔案副檔名為 *.txt
       
    Do While Len(nFile)
        If nFile <> "" Then
            File = Split(nFile, "_")
           
            'Macro通用
            If File(1) <> "" Then
                
                If File(1) = Left(OperationID, 2) Then
                    j = j + 1
                    If j = Down Then
                        MDIForm1.Text1.Text = ""             '清除內容
                        Open ("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\" & nFile) For Input As #2 '開啟文字檔
        
                        Line Input #2, str              '逐行讀取
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
                        MDIForm1.Text1.Text = ""             '清除內容
                        Open ("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\" & nFile) For Input As #2 '開啟文字檔
        
                        Line Input #2, str              '逐行讀取
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
                        MDIForm1.Text1.Text = ""             '清除內容
                        Open ("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\" & nFile) For Input As #2 '開啟文字檔
        
                        Line Input #2, str              '逐行讀取
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
                            MDIForm1.Text1.Text = ""             '清除內容
                            Open ("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\" & nFile) For Input As #2 '開啟文字檔
            
                            Line Input #2, str              '逐行讀取
                            b = "_"
                            message = Split(str, b)
                            For i = 0 To UBound(message)
                                MDIForm1.Text1.Text = MDIForm1.Text1.Text & message(i) & vbCrLf
                                MDIForm1.Text1.Font.Size = 13
                            Next i
            
                            Close #2
                        End If
                    End If
                    
                'Macro07條件
'                ElseIf File(1) = LineID Then
'                    j = j + 1
'                    If j = Down Then
'                        MDIForm1.Text1.Text = ""             '清除內容
'                        Open ("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\" & nFile) For Input As #2 '開啟文字檔
'
'                        Line Input #2, str              '逐行讀取
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
            
            


   
   MDIForm1.Label1 = Down & "/" & j & "  頁"
   'exProduct = ProductID
   
   GoTo finish
     


Err:
MsgBox "Fab槽連線失敗，請重新連線"

Exit Sub
finish:
exA_GlassID = GlassID


End Sub
