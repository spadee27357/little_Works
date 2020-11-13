Attribute VB_Name = "modSymbol"
Public Const CONFIG_NAME As String = "Config.ini"
Public Const RECIPEBODY_NAME As String = "RecipeBody.INI"
Public Const RECIPEBODYLIST_NAME As String = "RecipeList.INI"
'LogFile
Public m_Device As String

Public gAPLatestErrTime As String
Public gAPLatestErrMsg As String

'NET/H
Public Const g_nDeviceType = 24         'W register = 24

Public Const NETH_WORD_SIZE As Integer = 2

'For FAB5
Public Const NETH_EQREPORT_ADDR_START As Integer = &H1000
Public Const NETH_EQREPORT_ADDR_STARTS As String = "W1000"
Public Const NETH_EQREPORT_LENGTH As Integer = 256     '48     '<<52.14.303 dp87101 因為要抓EDC，所以範圍要擴大
Public Const NETH_MPLCINDICATE_ADDR_START As Integer = 2000 '&H0
Public Const NETH_MPLCINDICATE_ADDR_STARTS As String = "D2000" '"W0000"
Public Const NETH_MPLCINDICATE_LENGTH As Integer = 900 '896
'GroupNO
Public Const NETH_GROUPNO_ADDR_START As Integer = &H1001 '&HD01
Public Const NETH_GROUPNO_LENGTH As Integer = 1 '2
'RecipeNO
Public Const NETH_RECIPENO_ADDR_START As Integer = &H1002 '&HD02
Public Const NETH_RECIPENO_LENGTH As Integer = 1 '2
'GlassID
Public Const NETH_GLASSID_ADDR_START As Integer = &H182
Public Const NETH_GLASSID01_ADDR_START As Integer = 2032 '&H38
Public Const NETH_GLASSID02_ADDR_START As Integer = 2132 '&H106
Public Const NETH_GLASSID03_ADDR_START As Integer = 2232 '&H1C8
Public Const NETH_GLASSID04_ADDR_START As Integer = 2332 '&H225
Public Const NETH_GLASSID05_ADDR_START As Integer = 2432 '&H282
Public Const NETH_GLASSID06_ADDR_START As Integer = 2532 '&H2DF
Public Const NETH_GLASSID_LENGTH As Integer = 6 '10
'ProductID
Public Const NETH_PRODUCTID_ADDR_START As Integer = &H241
Public Const NETH_PRODUCTID01_ADDR_START As Integer = 2038 '&H3E
Public Const NETH_PRODUCTID02_ADDR_START As Integer = 2138 '&H10C
Public Const NETH_PRODUCTID03_ADDR_START As Integer = 2238 '&H1CE
Public Const NETH_PRODUCTID04_ADDR_START As Integer = 2338 '&H22B
Public Const NETH_PRODUCTID05_ADDR_START As Integer = 2438 '&H288
Public Const NETH_PRODUCTID06_ADDR_START As Integer = 2538 '&H2E5
Public Const NETH_PRODUCTID_LENGTH As Integer = 8 '14
'GroupID
Public Const NETH_GROUPID_ADDR_START As Integer = &H248
Public Const NETH_GROUPID01_ADDR_START As Integer = 2046 '&H46
Public Const NETH_GROUPID02_ADDR_START As Integer = 2146 '&H114
Public Const NETH_GROUPID03_ADDR_START As Integer = 2246 '&H1D6
Public Const NETH_GROUPID04_ADDR_START As Integer = 2346 '&H233
Public Const NETH_GROUPID05_ADDR_START As Integer = 2446 '&H290
Public Const NETH_GROUPID06_ADDR_START As Integer = 2546 '&H2ED
Public Const NETH_GROUPID_LENGTH As Integer = 10 '20

'<52.16.0410 dp87101 read CST ID
Public Const NETH_CSTID01_ADDR_START As Integer = 2056
Public Const NETH_CSTID02_ADDR_START As Integer = 2156
Public Const NETH_CSTID03_ADDR_START As Integer = 2256
Public Const NETH_CSTID04_ADDR_START As Integer = 2356
Public Const NETH_CSTID05_ADDR_START As Integer = 2456
Public Const NETH_CSTID06_ADDR_START As Integer = 2556
Public Const NETH_CSTID_LENGTH As Integer = 4
'52.16.0410>

'Public Const NETH_WORKNO_ADDR_START As Integer = &HD80
'Public Const NETH_WORK_LENGTH As Integer = 2

'Public Const NETH_JUDGERULE_ADDR_START As Integer = &H2C3
'Public Const NETH_JUDGERULE_LENGTH As Integer = 2
'OperatorID
Public Const NETH_OPERATORID_ADDR_START As Integer = &H10F8 '&H101D '&HD24  '<<35.14.303 dp87101 因機台W101D報空白，TK程式沒人改，先改抓EDC。
Public Const NETH_OPERATORID_LENGTH As Integer = 3 '8

'OperationID
Public Const NETH_OPERATIONID_ADDR_START As Integer = &H10F6  '放在MPLC W10F6 & W10F7 位置 所以會 = &H10F6 ，參照CIM設備規範書
Public Const NETH_OPERATIONID_LENGTH As Integer = 2 '長度2 F6和F7

'Current Time
Public Const NETH_CURRENT_TIME_ADDR_START As Integer = 2000 '&H0 '&H1
Public Const NETH_CURRENT_TIME_LENGTH As Integer = 3 '2
'Contrel GlassIn EQP, 20101101 !fms+
Public Const NETH_EQPCOUNT_ADDR_START As Integer = &H1006
Public Const NETH_EQPCOUNT_LENGTH As Integer = 1


Public g_bConnectStatus As Boolean

'Panel variable
Public Const GLASS_SIZE_X As Double = 1100000
Public Const GLASS_SIZE_Y As Double = 1300000
Public Const MAX_BLOCK_NUM As Integer = 6
Public Const MAX_PANEL_COL_NUM As Integer = 35
Public Const MAX_PANEL_ROW_NUM As Integer = 35
Public Const MAX_PANEL_NUM As Integer = MAX_PANEL_COL_NUM * MAX_PANEL_ROW_NUM '1225 '999
Public Const MAX_PANEL_DEFECT_NUM As Integer = 1600 '800 '500 '200
Public Const MAX_JUDGE_KIND As Integer = 4   'OK, NG, R, G_W, U_W       '<<52.13.1216 dp87101 add U_W
Public Const MAX_JUDGE_TYPE As Integer = 30
Public Const MAX_JUDGE_CODE As Integer = 30
Public Const MAX_CUT_NUM As Integer = 300 '100
Public Const CUT_SPLIT As Integer = 3 '2 '3

Public SYSTEM_RECIPE_PATH As String
Public SYSTEM_CAD_PATH  As String
Public SYSTEM_DEFECT_PATH As String

Public g_nBlockNum                     As Integer
Public g_nPanelNum                     As Integer
Public g_nCurPanelID                   As Integer
Public g_nDefectCount                  As Integer
Public g_bGlassExists                  As Boolean
Public g_sRecipeID                     As String
Public g_sGlassStartTime               As String
Public g_nCutNum                       As Integer
Public g_nGlsNRRuleNum                 As Integer   '20101118 !fms+
Public g_nGlsRCount                    As Integer   '20101118 !fms+
Public g_nGlsNCount                    As Integer   '20101118 !fms+
Public g_nGlsPCount                    As Integer   '20101118 !fms+
Public g_nGlsICount                    As Integer   '20101118 !fms+
Public g_nGlsGCount                    As Integer   '20101118 !fms+
Public g_nGlsNRRule2Num                As Integer   '20101118 !fms+
Public g_nGlsNRRule2Count              As Integer   '20101118 !fms+
Public g_sGlsNRJudge                   As String    '20101118 !fms+
Public g_nGlsWCount                    As Integer    '<<52.14.711 dp87101
Public g_nGlsUCount                    As Integer    '<<52.14.711 dp87101
Public g_nCutCol                       As Integer
Public g_nCutRow                       As Integer
Public g_sPanelInCut(MAX_CUT_NUM)      As String
Public g_nPanelNumInCut(MAX_CUT_NUM)   As String
Public g_nPanelNumFromCutFile          As Integer
Public g_nPanelNumFromDefectFile       As Integer
Public g_sOperatorID                   As String    '20120702 !fms*, For OPID Record (Fab5)
Public g_sOperationID                   As String
Public strLayoutInfo                    As String
Public OperationID                 As String
Public ProductID                   As String
Public GlassID                   As String
Public ex_GlassID                   As String
Public exC_GlassID                   As String
Public exA_GlassID                   As String
Public Recipe                      As Variant
Public ex_ProductID                As String
Public CoaterID                      As Variant
Public g_nGlassInEQP                   As Integer   'Real Glass In EQP
Public g_bMultPanelSelect              As Boolean   '20120419 !fms+, 多Panel選擇
Public g_sMultPanelLists()             As String    '20120419 !fms+, 多Panel選擇
Public g_sPanelJudgeFromCTCS           As String    '201310   efchen
Public g_nPanelNumFromCTCSFile         As String    '201310   efchen
Public g_sMsg                          As String    '<<52.14.804 dp87101

'20101101 !fms*, Move to Global
Public g_sGroupNo      As String
Public g_sProductID    As String
Public g_sGroupID      As String
Public g_sGlassID      As String
Public g_sPreGlassID   As String    '20101101 !fms+
Public g_sCurNetHTime  As String
Public g_sOldNetHTime  As String
Public g_sRecipeNo     As String
Public g_sPreGroupNo   As String    '20110503 !fms+
Public g_sPreProductID As String    '20110503 !fms+
Public g_sPreGroupID   As String    '20110503 !fms+
Public g_sPreRecipeNo  As String    '20110503 !fms+
Public g_sCSTID    As String            '<52.16.0410 dp87101 add CST ID infomation
'Public g_sJudgeRule    As String
'Public g_sWorkNo       As String   'output file for BC
Public g_sSendGlassID As String     '<52.19.2 dp87101

'20101115 !fms+, For Cut Panel Adjust
Public g_OrgSubFrameWidth As Single
Public g_OrgSubFrameHeight As Single
Public g_OrgSubPicxWidth As Single
Public g_OrgSubPicxHeight As Single
Public g_HasPIPanel As Boolean      '20111205 !fms+, For P,I Grade Panel

Public g_SystemInfo As TSystemInfo
Public g_PanelInfo() As TPanel
Public g_DefectInfo(MAX_PANEL_NUM - 1) As TDefect
Public g_SinglePanelCoord As TSinglePanelCoord
Public g_PanelType As TPanelType
Public g_RecipeBody As TRecipeBody

'/***************** Enum ***********************************/
Public Enum JUDGEKIND
   OK_JUDGE = 0
   NG_JUDGE = 1
   R_JUDGE = 2
   R_WJUDGE = 3
   R_UJUDGE = 4        '<<52.13.1216 dp87101 ADD
End Enum

'/***************** Type Structure *************************/

Public Type TSystemInfo
   m_sBC_C                 As String
   m_sBC_D                 As String
   m_sShop                 As String
   m_sLineID               As String
   '-------------------------------------
   m_nMacroNo              As Variant 'Integer 'Variant
   m_nCutMode              As Variant 'Integer '
   m_nSimulation           As Variant 'Integer '
   m_nPLCtype              As Variant 'Integer '
   m_nAutoMoveOut          As Variant 'Integer '
   m_nAutoMoveUpDate       As Variant 'Integer '
   m_nAutoMoveUpDateTime   As Variant 'Double
   m_nTopMode              As Variant 'Integer
   '-------------------------------------
   m_nTKNext               As Integer          '20110422 !fms+
   m_nTKNextIndex          As Integer          '20110422 !fms+
   'File Path
   m_sIniPath              As String
   m_sCutInfoPath          As String
   m_sCadFilePath          As String
   m_sDefectPath           As String
   m_LocalFilePath         As String           '20110421 !fms+
   m_TKNextPathName        As String           '20110422 !fms+
   
   'PLC-TCP Set
   m_sBCNo                      As String
   '-------------------------------------
   m_nLineNo                    As Variant 'Integer '
   m_nActCpuType                As Variant 'Integer '
   m_sActHostAddress            As String
   m_nActNetworkNumber          As Variant 'Integer '
   m_nActStationNumber          As Variant 'Integer '
   m_nActSourceNetworkNumber    As Variant 'Integer '
   m_nActSourceStationNumber    As Variant 'Integer '
   m_nActTimeOut                As Variant 'Integer '
    '-------------------------------------
   'Show Pre-Defect Code, 20120117
   'm_nShowPreDefect             As Integer
   'm_nShowDefectCodeNum         As Integer
   'm_nShowJudgeCode(3 - 1)      As Integer
End Type

Public Type TSinglePanelCoord
   m_nPanelID As Integer
   m_nLT_X As Double
   m_nLT_Y As Double
   m_nRD_X As Double
   m_nRD_Y As Double
End Type

Public Type TPanel
   m_nRowNum As Integer
   m_nColNum As Integer
   m_nPanelSizeX As Double
   m_nPanelSizeY As Double
   m_nPanelPitchX As Double
   m_nPanelPitchY As Double
   m_Layout() As TSinglePanelCoord
End Type

Public Type TJudgeInfo
   m_sCode As String
   m_sDesc As String
   m_sName As String
End Type

Public Type TJudgeType  '以模為單位, 讀檔用
   m_sType As String
   m_Info() As TJudgeInfo
End Type

Public Type TJudge  '讀檔用
   m_sKind As String   'NG, R, OK
   m_Type() As TJudgeType
End Type

Public Type TJudgeOutput
   m_sJudge As String   'NG, R, OK
   m_sType  As String
   m_sCode  As String   '641, 642, 645 and so on
   m_sDesc  As String
   m_sName  As String
End Type

Public Type TDefectInfo 'max is 200
   m_sName As String    ' keep original judge name
   m_sCode As String    ' keep original judge code
   m_nX As Double
   m_nY As Double
   m_sLine1 As String
   m_sLine2 As String
End Type

Public Type TDefect  '以模為單位
   m_bOX As Boolean
   m_btokki As Boolean       '20101130 !fms+
   m_bEverJudge As Boolean
   m_sPanelHeader As String
   m_nDefectNum As Integer
   m_Judge As TJudgeOutput
   m_DefectInfo(MAX_PANEL_DEFECT_NUM) As TDefectInfo
End Type

Public Type TRecipeField
   m_nPanelStartPosX As Double
   m_nPanelStartPosY As Double
   m_nPanelSizeX     As Double
   m_nPanelSizeY     As Double
   m_nPanelPitchX    As Double
   m_nPanelPitchY    As Double
   m_nPanelRowNum    As Double
   m_nPanelColNum    As Double
   m_nFrameU         As Double
   m_nFrameD         As Double
   m_nFrameL         As Double
   m_nFrameR         As Double
End Type

Public Type TRecipeBody
   m_sRecipeID       As String
   m_sAlignerLOPID   As String
   m_sAlignerROPID   As String
   m_sCADFileName    As String
   m_sCoaterLOPID    As String
   m_sCoaterROPID    As String
   m_nItem2LoadFlag  As Integer
   m_nJudgeRule      As Integer
   m_sMacroOPID      As String
   m_nOPTotal        As Integer
   m_sOPID(12 - 1)   As String
   m_sRecipeVer      As String
   m_nRepairJudge    As Integer
End Type

Public Type TPanelType
   m_sTypeName(MAX_PANEL_NUM - 1) As String
   m_nTypeNum(MAX_PANEL_NUM - 1) As Integer
End Type

'CPU Type 對照表
' Q系列CPU
' CPU_Q02CPU = &H22                                          ' Q02(H) Q
' CPU_Q06CPU = &H23                                          ' Q06H   Q
' CPU_Q12CPU = &H24                                          ' Q12H   Q
' CPU_Q25CPU = &H25                                          ' Q25H   Q
' CPU_Q00JCPU = &H30                                         ' Q00J   Q
' CPU_Q00CPU = &H31                                          ' Q00    Q
' CPU_Q01CPU = &H32                                          ' Q01    Q
' CPU_Q12PHCPU = &H41                                        ' Q12PHCPU Q
' CPU_Q25PHCPU = &H42                                        ' Q25PHCPU Q
' CPU_Q12PRHCPU = &H43                                       ' Q12PRHCPU Q
' CPU_Q25PRHCPU = &H44                                       ' Q25PRHCPU Q
' CPU_Q25SSCPU = &H55                                        ' Q25SS
'隨著讀取plc記憶體的大小需做大小的調整
Public MPLCdataW1(1000) As Long
Public MPLCdataW2(1000) As Long
Public MPLCdataB(1000) As Long

