Attribute VB_Name = "抓取local資料"
Sub localdata()
On Error GoTo ErrorHandle
    Dim fso As FileSystemObject
    Dim fid As TextStream
    Dim read As String, aFile As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Oper_1 = "" Then
        GoTo Number2
    ElseIf Oper_1 <> "" Then
        aFile = Dir("\\10.91.1.83\main$\RTQCS\" & Oper_1 & ".txt")
        If aFile = "" Then
            dFile = Dir("\\10.91.1.83\DMAC$\DMRV\" & Oper_1 & ".txt")
            If dFile = "" Then
                GoTo Number2
            End If
        End If
    End If
    
    aFile = Dir("\\10.91.1.83\main$\RTQCS\" & Oper_1 & ".txt")
    If aFile <> "" Then
        RTQCS = FileDateTime("\\10.91.1.83\main$\RTQCS\" & Oper_1 & ".txt")
    Else
        RTQCS = ""
    End If
    
    dFile = Dir("\\10.91.1.83\DMAC$\DMRV\" & Oper_1 & ".txt")
    If dFile <> "" Then
        DMRV = FileDateTime("\\10.91.1.83\DMAC$\DMRV\" & Oper_1 & ".txt")
    Else
        DMRV = ""
    End If
    
    If RTQCS > DMRV Or DMRV = "" Then
        Choose_which1 = 1
    ElseIf DMRV > RTQCS Or RTQCS = "" Then
        Choose_which1 = 2
    End If
    
    If Choose_which1 = 1 Then
        Open "\\10.91.1.83\main$\RTQCS\" & Oper_1 & ".txt" For Input As #1
    ElseIf Choose_which1 = 2 Then
        Open "\\10.91.1.83\DMAC$\DMRV\" & Oper_1 & ".txt" For Input As #1
    End If
    
    Do While Not EOF(1)
        Line Input #1, read
        nFile = Split(read, ",")
        aFile = read
        Product_1 = nFile(0)
        Operation_1 = nFile(1)
        Gls_1 = nFile(2)
        Coater_1 = nFile(3)
        'Aligner_1 = nFile(4)
    Loop
    Close #1
    
    MDIForm1.Label1 = "產品: " & Product_1
    MDIForm1.Label2 = "站點: " & Operation_1
    MDIForm1.Label3 = "Gls_ID: " & Gls_1
    MDIForm1.Label4 = "Coater: " & Coater_1
    MDIForm1.Label5 = "Aligner: " & Aligner_1
    MDIForm1.Label1.Font.Size = 13
    MDIForm1.Label2.Font.Size = 13
    MDIForm1.Label3.Font.Size = 13
    MDIForm1.Label4.Font.Size = 13
    MDIForm1.Label5.Font.Size = 13
    
Number2:

    If Oper_2 = "" Then
        GoTo finish
    ElseIf Oper_2 <> "" Then
        aFile = Dir("\\10.91.1.83\main$\RTQCS\" & Oper_2 & ".txt")
        If aFile = "" Then
            dFile = Dir("\\10.91.1.83\DMAC$\DMRV\" & Oper_2 & ".txt")
            If dFile = "" Then
                GoTo finish
            End If
        End If
    End If

    aFile = Dir("\\10.91.1.83\main$\RTQCS\" & Oper_2 & ".txt")
    If aFile <> "" Then
        RTQCS = FileDateTime("\\10.91.1.83\main$\RTQCS\" & Oper_2 & ".txt")
    Else
        RTQCS = ""
    End If
    
    dFile = Dir("\\10.91.1.83\DMAC$\DMRV\" & Oper_2 & ".txt")
    If dFile <> "" Then
        DMRV = FileDateTime("\\10.91.1.83\DMAC$\DMRV\" & Oper_2 & ".txt")
    Else
        DMRV = ""
    End If
    
    If RTQCS > DMRV Or DMRV = "" Then
        Choose_which2 = 1
    ElseIf DMRV > RTQCS Or RTQCS = "" Then
        Choose_which2 = 2
    End If
    a = a
    If Choose_which2 = 1 Then
        Open "\\10.91.1.83\main$\RTQCS\" & Oper_2 & ".txt" For Input As #2
    ElseIf Choose_which2 = 2 Then
        Open "\\10.91.1.83\DMAC$\DMRV\" & Oper_2 & ".txt" For Input As #2
    End If
    
    Do While Not EOF(2)
        Line Input #2, read
        nFile = Split(read, ",")
        aFile = read
        Product_2 = nFile(0)
        Operation_2 = nFile(1)
        Gls_2 = nFile(2)
        Coater_2 = nFile(3)
        'Aligner_2 = nFile(4)

    Loop
    Close #2
    
    MDIForm1.Label7 = "產品: " & Product_2
    MDIForm1.Label8 = "站點: " & Operation_2
    MDIForm1.Label9 = "Gls_ID: " & Gls_2
    MDIForm1.Label10 = "Coater: " & Coater_2
    MDIForm1.Label11 = "Aligner: " & Aligner_2
    MDIForm1.Label7.Font.Size = 13
    MDIForm1.Label8.Font.Size = 13
    MDIForm1.Label9.Font.Size = 13
    MDIForm1.Label10.Font.Size = 13
    MDIForm1.Label11.Font.Size = 13


finish:
Exit Sub
ErrorHandle:
MsgBox "抓取local資料錯誤"
End Sub
