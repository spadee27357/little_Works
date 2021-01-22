Attribute VB_Name = "辨識"
Private Declare Function PrintWindow Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long, ByVal nFlags As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020




Sub Look()
 On Error Resume Next
 

'--------------------------------------------------------------------------------------------------

        Dim lngDesktopHwnd As Long
        Dim lngDesktopDC As Long

        lngDesktopHwnd = GetDesktopWindow
        lngDesktopDC = GetDC(lngDesktopHwnd)


        
        
        'Call BitBlt(MDIForm1.Picture2.hdc, 0, 0, 300, 40, lngDesktopDC, 530, 40, SRCCOPY)
        'MDIForm1.Picture2.Picture = MDIForm1.Picture2.Image
        'SavePicture MDIForm1.Picture2.Image, "I:\\製造部\組長專用\顏伯丞\圖片\Macro07\站點.jpg"
        'Call ReleaseDC(lngDesktopHwnd, lngDesktopDC)

a = a
'--------------------------------------------------------------------------------------------------
 

    Dim strLPN As String     '初始化並加載文檔

    Set miDoc = CreateObject("MODI.Document")            '創建對象
    miDoc.Create "I:\\製造部\組長專用\顏伯丞\圖片\無提示\test2.jpg"                         '加載圖片文件
    Screen.MousePointer = vbHourglass                    '設置光標忙
    miDoc.Images(0).OCR miLANG_ENGLISH, True, True '有用的就此一句
    Set modiLayout = miDoc.Images(0).Layout              '讀出數據
    strLayoutInfo = modiLayout.Text
    
    'MsgBox strLayoutInfo, vbInformation + vbOKOnly, "Layout Information"
    MDIForm1.Label3 = "Operation ID: " & Left(strLayoutInfo, 4)  ', vbInformation + vbOKOnly,
    
    'If Dir("I:\製造部\組長專用\顏伯丞\圖片\Macro07") = "" Then MkDir ("I:\製造部\組長專用\顏伯丞\圖片\Macro07")
  
    '匯出TXT
    Dim MYstr As String, i As Integer                '定義屬性
    Open "I:\製造部\組長專用\顏伯丞\圖片\Macro07\站點.txt" For Output As #1     '定義Output File位置
             '輸出的內容
    MYstr = Left(strLayoutInfo, 4)
    Print #1, MYstr
    Close #1
    
    
    a = a
    Set modiLayout = Nothing
    Set miDoc = Nothing
    Screen.MousePointer = vbDefault
    
    

End Sub
