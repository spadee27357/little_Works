Attribute VB_Name = "抓取local資料"
Sub localdata()
    'On Error Resume Next
    'C:\R1378\MMI\MMI_INI\CurGlassInfo.INI
    a = a
    Dim FSO As FileSystemObject
    Dim fid As TextStream
    Dim Read As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")

    FSO.CopyFile "C:\R1378\MMI\MMI_INI\CurGlassInfo.INI", "D:\LogFile\MACRO RUN\local data\", True       'SAMP 'Final
    Open "D:\LogFile\MACRO RUN\local data\CurGlassInfo.INI" For Input As #1
    Do While Not EOF(1)

        'CurCoaterID =
        Input #1, Read
        j = InStr(Read, "CurCoaterID = ")
        If j = 1 Then
            Coater = Split(Read, "CurCoaterID = ")
            If Coater(1) = """""" Then
                OperationID = "noon"
                Recipe = "noon"
                ProductID = ""
                GlassID = ""
                CoaterID = ""
                Close #1
                GoTo finish
            End If

            CoaterID = Mid(Coater(1), 8, 2)
            a = a
            Read = ""
            j = ""
            'CurGlassID =
            Do While Not EOF(1)
                Input #1, Read
                j = InStr(Read, "CurGlassID = ")
                'MsgBox 2
                If j = 1 Then
                    GlassA = Split(Read, "CurGlassID = ")
                    GlassID = Mid(GlassA(1), 2, 10)
                    Read = ""
                    j = ""
                    Exit Do
                End If
            Loop
                    'CurOperID =
                    Do While Not EOF(1)
                        Input #1, Read
                        j = InStr(Read, "CurOperID = ")
                        If j = 1 Then
                            OperID = Split(Read, "CurOperID = ")
                            Recipe = "0" & OperID(1)
                            Recipe = Right(Recipe, 4)
                            j = ""
                            a = a
                            Exit Do
                        End If
                    Loop
                        'CurProductID =
                        Do While Not EOF(1)
                            Input #1, Read
                            j = InStr(Read, "CurProductID = ")
                            If j = 1 Then
                                Product = Split(Read, "CurProductID = ")
                                ProductID = Mid(Product(1), 2, 10)
                                Read = ""
                                j = ""
                                GoTo RecipeBody
                            End If
                        Loop


        End If
    Loop
    

RecipeBody:

    Close #1
    a = a

    Open "C:\R1378\MMI\MMI_INI\RecipeBody.ini" For Input As #2               'Final
    Do While Not EOF(2)
        Input #2, Read
        j = InStr(Read, "[Recipe" & Recipe & "]")
        If j = 1 Then
            j = ""
            Do While Not EOF(2)
            Input #2, Read
            k = InStr(Read, "Macro Operation ID = ")
                If k = 1 Then
                    Id = Split(Read, "Macro Operation ID = ")
                    OperationID = Mid(Id(1), 2, 4)
                    Read = ""
                    j = ""
                    Close #2
                    GoTo Map
                End If
            Loop
            
        End If
    Loop


Map:

a = a
If ex_GlassID = GlassID Then
    GoTo finish
End If




'
' Set xlApp = CreateObject("Excel.Application")
'
' Set xlBook = xlApp.Workbooks.Open("C:\Macro\Macro Run\(機台)SAMP MACRO JPG.xls")
'     xlApp.DisplayAlerts = False
'     xlApp.Visible = False
'     xlBook.Activate
'     Set xlSheet = xlBook.Worksheets("Map")
'     xlSheet.Activate
'     xlSheet.cells(2, 18) = ProductID
'     xlSheet.cells(2, 20) = GlassID
'     xlSheet.cells(2, 22) = OperationID
'     xlSheet.cells(2, 24) = CoaterID
'     a = a
'     xlApp.Run ("automatic")
'     xlApp.Quit

'------------------------------------------------------------------------------------------
a = a
'nFile = Dir("D:\LogFile\MACRO RUN\" & ex_ProductID & "\" & ex_ProductID & "_1-" & ex_GlassID & "*.jpg")
'If nFile <> "" Then
'    File = Split(nFile, "_")
'    If File(1) = "1-" & ex_GlassID & "無缺陷.jpg" Then
'        Kill "D:\LogFile\MACRO RUN\" & ex_ProductID & "\" & ex_ProductID & "_1-" & ex_GlassID & "無缺陷.jpg"
'        Kill "D:\LogFile\MACRO RUN\" & ex_ProductID & "\" & ex_ProductID & "_1-" & ex_GlassID & "無缺陷.txt"
'    End If
'End If

     ex_GlassID = GlassID
     ex_ProductID = ProductID
     




finish:


a = a
End Sub


Private Sub Delay(ASecond As Integer)
    Dim before
    before = Timer
    Do
    DoEvents
    Loop Until (Int(Timer - before) = ASecond)
End Sub

