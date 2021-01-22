Attribute VB_Name = "modLog_TrcTrx"
Option Explicit

Public Enum eLevel
    Error = 1
    Warn = 2
    Trace = 3
    block = 4
    Dbg = 5
    tcp = 6
    EQP = 7
    Glass = 8
    
End Enum

Public G_TimeLog As String

'Level -> 1:error 2:warning 3:trace 4:debug 5:debug2
Public Function TraceOut(ByVal msg As String, Level As eLevel) As Boolean
1        On Error Resume Next
    
2        'If Level < G_BC.mTraceLevel Then
3        Select Case Level
             Case Error
5                msg = "[ Error ] " & msg
6            Case Warn
7                msg = "[ Warn  ] " & msg
8            Case Trace
9                msg = "[ Trace ] " & msg
10           Case Dbg
11               msg = "[ Dbg1  ] " & msg
12           Case block
13               msg = "[ Block ] " & msg
14           Case tcp
15               msg = "[ Tcp ] " & msg
16       End Select
    
18       UT_ListMsgLog msg, Level
    
20       If Level = Error Then
21           gAPLatestErrTime = Format(Now, "YYYY/MM/DD HH:MM:SS")
22           gAPLatestErrMsg = msg
23       End If
  
24       TraceOut = SaveTrcLog(msg)
         
25       msg = "" 'stress
    
End Function
Public Function SaveTcpLog(msg As String) As Boolean
1        Dim sFolder As String
2        Dim FileNum As Integer
3        Dim strTime As String
4        On Error GoTo ErrSaveTRCLog
5        SaveTcpLog = False
6        strTime = Format(Now, "YYYY/MM/DD hh:mm:ss")
7        '**Check Folder is exist ***********************************
8        sFolder = m_Device & "\" & UCase(App.EXEName) & "\" & Format(Now, "yyyymmdd")
9        If Dir(sFolder, vbDirectory) = "" Then
10           Call NewCreateFolder(sFolder)
11       End If
12       '**********************************************************
    
14       'Open File Save Msg   *************************************
15       FileNum = FreeFile()
16       sFolder = sFolder & "\Tcp" & Format(strTime, "hh") & ".Log"         '20031007 lee modify
17       Open sFolder For Append As #FileNum
18       Print #FileNum, strTime & " " & msg
19       Close #FileNum
20       msg = "" 'stress
21       SaveTcpLog = True
22       Exit Function
23 ErrSaveTRCLog:
24       Close #FileNum
25       err.Clear
End Function

Public Function SaveEqpLog(msg As String) As Boolean
1        Dim sFolder As String
2        Dim FileNum As Integer
3        Dim strTime As String
4        On Error GoTo ErrSaveTRCLog
5        SaveEqpLog = False
6        strTime = Format(Now, "YYYY/MM/DD hh:mm:ss")
7        '**Check Folder is exist ***********************************
8        sFolder = m_Device & "\" & UCase(App.EXEName) & "\" & Format(Now, "yyyymmdd")
9        If Dir(sFolder, vbDirectory) = "" Then
10           Call NewCreateFolder(sFolder)
11       End If
12       '**********************************************************
    
14       'Open File Save Msg   *************************************
15       FileNum = FreeFile()
16       sFolder = sFolder & "\Eqp" & Format(strTime, "hh") & ".Log"         '20031007 lee modify
17       Open sFolder For Append As #FileNum
18       Print #FileNum, strTime & " " & msg
19       Close #FileNum
20       msg = "" 'stress
21       SaveEqpLog = True
22       Exit Function
23 ErrSaveTRCLog:
24       Close #FileNum
25       err.Clear
End Function

Public Function SaveGlassLog(msg As String) As Boolean
1        Dim sFolder As String
2        Dim FileNum As Integer
3        Dim strTime As String
4        On Error GoTo ErrSaveTRCLog
5        SaveGlassLog = False
6        strTime = Format(Now, "YYYY/MM/DD hh:mm:ss")
7        '**Check Folder is exist ***********************************
8        sFolder = m_Device & "\" & UCase(App.EXEName) & "\" & Format(Now, "yyyymmdd")
9        If Dir(sFolder, vbDirectory) = "" Then
10           Call NewCreateFolder(sFolder)
11       End If
12       '**********************************************************
    
14       'Open File Save Msg   *************************************
15       FileNum = FreeFile()
16       sFolder = sFolder & "\Glass" & Format(strTime, "hh") & ".Log"         '20031007 lee modify
17       Open sFolder For Append As #FileNum
18       Print #FileNum, strTime & " " & msg
19       Close #FileNum
20       msg = "" 'stress
21       SaveGlassLog = True
22       Exit Function
23 ErrSaveTRCLog:
24       Close #FileNum
25       err.Clear
End Function
Public Function SaveTrcLog(msg As String) As Boolean
1        Dim sFolder As String
2        Dim FileNum As Integer
3        Dim strTime As String
4        On Error GoTo ErrSaveTRCLog
5        SaveTrcLog = False
6        strTime = Format(Now, "YYYY/MM/DD hh:mm:ss")
    
8        '**Check Folder is exist ***********************************
9        m_Device = IIf(Right(m_Device, 1) = "\", m_Device, m_Device & "\")
10       sFolder = m_Device & UCase(App.EXEName) & "\" & Format(Now, "yyyymmdd")
11       If Dir(sFolder, vbDirectory) = "" Then
12           Call NewCreateFolder(sFolder)
13       End If
14       '**********************************************************
    
16       'Open File Save Msg   *************************************
17       FileNum = FreeFile()
18       sFolder = sFolder & "\Trc" & Format(strTime, "hh") & ".Log"         '20031007 lee modify
19       'sFolder = sFolder & "\Trc-2010-10-28.Log"
20       Open sFolder For Append As #FileNum
21       Print #FileNum, strTime & " " & msg
22       Close #FileNum
23       'Open App.Path & G_WORK_FILE For Random As #intFileHandle Len = Len(T_WorkF)
24       'Put #intFileHandle, delRecNo, T_WorkF    ' 將資料清空記錄寫入檔案中。
25       'Close #intFileHandle
26       msg = "" 'stress
27       SaveTrcLog = True
28       Exit Function
29 ErrSaveTRCLog:
30       Close #FileNum
31       err.Clear
End Function

Public Function NewCreateFolder(FoldSpec As String) As Boolean
1        On Error Resume Next
2        Dim G_FSO As New clsFSO
3        If Right(FoldSpec, 1) <> "\" Then FoldSpec = FoldSpec & "\"
4        NewCreateFolder = G_FSO.Create_Dir_Struct(FoldSpec)
    
End Function

Private Function UT_ListMsgLog(ByVal newLog As String, ByVal newLogType As eLevel)
1        Dim lstObj As ListView
2        Dim intTemp As Integer
    
    
5        On Error Resume Next
6        '-------------將Log檔寫入
    
8        'DoEvents
    
10       If newLogType = Error Then
11           ''Set lstObj = frmMain.lstErr
12       Else
13           ''Set lstObj = frmMain.lstTrx
14       End If
    '------------------------
16       lstObj.ListItems.Add , , "[" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "]" & newLog
    
18       If lstObj.ListItems.Count > 100 Then
19           'For intTEmp = 1 To 90
20           '    lstObj.ListItems.Remove 1
21           'Next
22           lstObj.ListItems.Remove 1
23       End If
    
End Function

