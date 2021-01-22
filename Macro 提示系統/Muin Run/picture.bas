Attribute VB_Name = "圖片"
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
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\CV Macro\" & ProductID & "\*.jpg")
    
    'MUIN 20
    'nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\*.jpg")  '設定想要處理的目錄為 C:\123, 處理的檔案副檔名為 *.txt

a = a
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
'    nFile = Dir("D:\LogFile\MACRO RUN\" & ProductID & "\*.jpg")
'        Do While Len(nFile)
'    If nFile <> "" Then
'            File = Split(nFile, "_")
'            'If File(2) = "" Or File(1) = Left(OperationID, 2) Or File(1) = "07" Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then 'Macro07條件
'            If File(2) = "" Or File(1) = Left(OperationID, 2) Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then 'Macro通用
'            'If CoaterID = "12" And File(1) = "R" Then  'MUIN通用
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
'--------------------------------D槽

'--------------------------------臨時加強檢
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\CV Macro\*.txt")
    Do While Len(nFile) '計算資料夾內有幾張圖片
a = a
        If nFile <> "" Then
            aFile = Split(nFile, ".")
   
            'Macro通用
            If aFile(0) <> "" Then
        
                j = j + 1
                If j = Down Then
                    MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\CV Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".jpg")
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
'--------------------------------臨時加強檢

    'MUIN
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\CV Macro\" & ProductID & "\*.jpg")
    
    'MUIN 20
    'nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\*.jpg")
    Do While Len(nFile) '計算資料夾內有幾張圖片
a = a
        If nFile <> "" Then
            File = Split(nFile, "_")

            If (CoaterID = "13" And File(1) = "L") Or (CoaterID = "12" And File(1) = "R") Or File(1) = "Coater" Then
                j = j + 1
                If j = Down Then
                    'MUIN
                    'MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\CV Macro\" & ProductID & "\" & nFile)
                
                    'MUIN 20
                    MDIForm1.Image1.Picture = LoadPicture("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\CV Macro\" & ProductID & "\" & nFile)
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
'--------------------------------D槽
'     Do While Len(nFile)
'nFile = Dir("D:\LogFile\MACRO RUN\" & ProductID & "\*.txt")
'
'
'        If nFile <> "" Then
'            File = Split(nFile, "_")
'            'If File(2) = "" Or File(1) = Left(OperationID, 2) Or File(1) = "07" Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then  'Macro07條件
'            If File(2) = "" Or File(1) = Left(OperationID, 2) Or (CoaterID = "12" And File(1) = "R") Or (CoaterID = "13" And File(1) = "L") Then 'Macro通用
'            a = a
'            'If CoaterID = "12" And File(1) = "R" Then     'MUIN通用
'                j = j + 1
'                If j = Down Then
'                    MDIForm1.Text1.Text = ""             '清除內容
'                    Open ("D:\LogFile\MACRO RUN\" & ProductID & "\" & nFile) For Input As #1  '開啟文字檔
'
'                    Line Input #1, str              '逐行讀取
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
'--------------------------------D槽

'--------------------------------臨時加強檢
    nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\CV Macro\*.txt")
    Do While Len(nFile) '計算資料夾內有幾張圖片
a = a
        If nFile <> "" Then
            aFile = Split(nFile, ".")
            
            '通用
            If aFile(0) <> "" Then
        
                j = j + 1
                If j = Down Then
                    MDIForm1.Text1.Text = ""
                    Open ("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\CV Macro\" & aFile(0) & "\" & ProductID & "\" & ProductID & "_" & aFile(0) & ".txt") For Input As #3
                    
                    Line Input #3, str              '逐行讀取
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
'--------------------------------臨時加強檢

'MUIN
nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\CV Macro\" & ProductID & "\*.txt")

'MUIN 20
'nFile = Dir("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\SAMP Macro\" & ProductID & "\*.txt")  '設定想要處理的目錄為 C:\123, 處理的檔案副檔名為 *.txt

    Do While Len(nFile)
        If nFile <> "" Then
            File = Split(nFile, "_")

            If (CoaterID = "13" And File(1) = "L") Or (CoaterID = "12" And File(1) = "R") Or File(1) = "Coater" Then
                j = j + 1
                If j = Down Then
                    MDIForm1.Text1.Text = ""             '清除內容
                    
                    'MUIN
                    'Open ("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\CV Macro\" & ProductID & "\" & nFile) For Input As #1 '開啟文字檔
                    
                    'MUIN 20
                    Open ("\\10.91.40.40\fabsh$\cf5\製造部\組長專用\顏伯丞\圖片\CV Macro\" & ProductID & "\" & nFile) For Input As #1 '開啟文字檔

                    Line Input #1, str              '逐行讀取
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

   MDIForm1.Label1 = Down & "/" & j & "  頁"

   
    
     
finish:
exA_GlassID = GlassID
End Sub
