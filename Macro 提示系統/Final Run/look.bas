Attribute VB_Name = "����"
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
        'SavePicture MDIForm1.Picture2.Image, "I:\\�s�y��\�ժ��M��\�C�B��\�Ϥ�\Macro07\���I.jpg"
        'Call ReleaseDC(lngDesktopHwnd, lngDesktopDC)

a = a
'--------------------------------------------------------------------------------------------------
 

    Dim strLPN As String     '��l�ƨå[������

    Set miDoc = CreateObject("MODI.Document")            '�Ыع�H
    miDoc.Create "I:\\�s�y��\�ժ��M��\�C�B��\�Ϥ�\�L����\test2.jpg"                         '�[���Ϥ����
    Screen.MousePointer = vbHourglass                    '�]�m���Ц�
    miDoc.Images(0).OCR miLANG_ENGLISH, True, True '���Ϊ��N���@�y
    Set modiLayout = miDoc.Images(0).Layout              'Ū�X�ƾ�
    strLayoutInfo = modiLayout.Text
    
    'MsgBox strLayoutInfo, vbInformation + vbOKOnly, "Layout Information"
    MDIForm1.Label3 = "Operation ID: " & Left(strLayoutInfo, 4)  ', vbInformation + vbOKOnly,
    
    'If Dir("I:\�s�y��\�ժ��M��\�C�B��\�Ϥ�\Macro07") = "" Then MkDir ("I:\�s�y��\�ժ��M��\�C�B��\�Ϥ�\Macro07")
  
    '�ץXTXT
    Dim MYstr As String, i As Integer                '�w�q�ݩ�
    Open "I:\�s�y��\�ժ��M��\�C�B��\�Ϥ�\Macro07\���I.txt" For Output As #1     '�w�qOutput File��m
             '��X�����e
    MYstr = Left(strLayoutInfo, 4)
    Print #1, MYstr
    Close #1
    
    
    a = a
    Set modiLayout = Nothing
    Set miDoc = Nothing
    Screen.MousePointer = vbDefault
    
    

End Sub
