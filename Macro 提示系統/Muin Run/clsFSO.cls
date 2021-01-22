VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ***************************************************************************
' Module:        clsFSO
'
' Description:   Module with several routines to make use of the
'                Scripting.fileSystemObject.  Be sure to make a reference
'                to "Microsoft Scripting Runtime" (scrrun.dll) in your
'                project.
'
'                Visit http://msdn.microsoft.com/default.asp and type
'                "FileSystemObject" (without the quotes) in the search box
'                to get additional information.
'
'                To get a hard copy of the VBScript documentation or an updated
'                version, visit:
'
'    http://msdn.microsoft.com/scripting/default.htm?/scripting/vbscript/doc/default.htm
'
'                and click on "Downloads" on the right side.
'                o  Click on the first hot link at the top of the page to get
'                   the latest version of the scripting engine.  Scroll to
'                   the bottom of the page.
'                o  Click on the "32-bit Vbscript ..." hot link to get the
'                   documentation.
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-NOV-2000  Kenneth Ives  kenaso@home.com
'              Routine created
' 24-MAR-2001  Kenneth Ives  kenaso@home.com
'              Added routines and documentation
' ***************************************************************************

' ---------------------------------------------------------------------------
' MISCELLANEOUS CONSTANTS
' ---------------------------------------------------------------------------
Private Const BIF_RETURNONLYFSDIRS As Long = 1
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const SEE_MASK_INVOKEIDLIST As Long = &HC
Private Const SEE_MASK_NOCLOSEPROCESS As Long = &H40
Private Const SEE_MASK_FLAG_NO_UI As Long = &H400
Private Const OF_READWRITE As Long = &H2
Private Const MAX_SIZE As Long = 260
Private Const MAX_FILE_COUNT As Long = 32768

' ---------------------------------------------------------------------------
' Store drive information
' ---------------------------------------------------------------------------
Private Type DriveInfo
    DriveLetter As String
    ShareName As String
    Path As String
    RootFolder As String
    SerialNumber As String
    VolumeName As String
    TotalSize As String
    AvailableSpace As String
    FreeSpace As String
    DriveType As String
    FileSystem As String
    IsReady As String
End Type

' ---------------------------------------------------------------------------
' Store file information
' ---------------------------------------------------------------------------
Private Type FileInfo
    Drive As String
    Path As String
    ParentFolder As String
    Size As String
    Type As String
        Extension As String
        DateLastAccessed As String
        DateLastModified As String
        DateCreated As String
        Attributes As String
    End Type
    
    ' ---------------------------------------------------------------------------
    ' Required for SHBrowseForFolder API call
    ' ---------------------------------------------------------------------------
    Private Type BrowseInfo
        hWndOwner As Long
        pIDLRoot As Long
        pszDisplayName As Long
        lpszTitle As Long
        ulFlags As Long
        lpfnCallback As Long
        lParam As Long
        iImage As Long
    End Type
    
    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
    
    Private Declare Function SHBrowseForFolder Lib "shell32" _
            (lpbi As BrowseInfo) As Long
    
    Private Declare Function SHGetPathFromIDList Lib "shell32" _
            (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
            (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    
    ' ---------------------------------------------------------------------------
    ' Needed for a list of all available drive letters
    ' ---------------------------------------------------------------------------
    Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
            Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
            ByVal lpBuffer As String) As Long
    
    ' ---------------------------------------------------------------------------
    ' Needed to reference/create temporary folders/files
    ' ---------------------------------------------------------------------------
    Private Declare Function GetTempPath Lib "kernel32.dll" _
            Alias "GetTempPathA" (ByVal nBufferLength As Long, _
            ByVal lpBuffer As String) As Long
    
    Private Declare Function GetTempFileName Lib "kernel32.dll" _
            Alias "GetTempFileNameA" (ByVal lpszPath As String, _
            ByVal lpPrefixString As String, ByVal wUnique As Long, _
            ByVal lpTempFileName As String) As Long
    
    ' ---------------------------------------------------------------------------
    ' Needed to determine space requirements of a drive
    ' ---------------------------------------------------------------------------
    ' Do we have the ability to use GetDiskFreeSpaceEx
    Private Declare Function LoadLibrary Lib "kernel32" _
            Alias "LoadLibraryA" _
            (ByVal lpLibFileName As String) As Long
    
    ' Verify the proceedure address within the kernel32 dll
    Private Declare Function GetProcAddress Lib "kernel32" _
            (ByVal hModule As Long, ByVal lpProcName As String) As Long
    
    ' Decrement the DLL counter when we are finished.  This is
    ' our safety net.
    Private Declare Function FreeLibrary Lib "kernel32" _
            (ByVal hLibModule As Long) As Long
    
    ' Drives over 2GB (2,147,483,647 bytes)
    Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
            Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
            FreeBytesAvailableToCaller As Currency, _
            TotalNumberOfBytes As Currency, _
            TotalNumberOfFreeBytes As Currency) As Long
    
    ' Drives under 2GB (2,147,483,647 bytes)
    Private Declare Function GetDiskFreeSpace Lib "kernel32" _
            Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
            lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
            lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
    
    ' ---------------------------------------------------------------------------
    ' Define module level variables
    ' ---------------------------------------------------------------------------
    Private FI As FileInfo
    Private DI As DriveInfo
    Private m_dblCount As Double
    
    ' ---------------------------------------------------------------------------
    ' Make a reference to "Microsoft Scripting Runtime"  (scrrun.dll)
    ' ---------------------------------------------------------------------------
    Private FSO As Scripting.FileSystemObject
    
    ' ---------------------------------------------------------------------------
    ' Used by an external click event to stop processing when doing something in
    ' a loop, like gathering a list of files.
    '
    ' Example code in form cmdStop_Click event
    '
    '         g_blnStopProcessing = True
    '         cFSO.CancelProcessing = True
    '
    ' The class module looks for CancelProcessing while the external code looks
    ' for a global variable named g_blnStopProcessing.  To get this to work, I
    ' had to make cFSO a global variable.
    ' ---------------------------------------------------------------------------
    Public CancelProcessing As Boolean
    
    ' ---------------------------------------------------------------------------
    ' This event used to display the file count as they are located.
    ' See Get_FileList() routine.
    ' ---------------------------------------------------------------------------
    Public Event CountFiles(dblCount As Double, strPath As String)
    
    ' ---------------------------------------------------------------------------
    ' This event used to display the accumulation count as it is incremented.
    ' See Get_Totals() routine.
    ' ---------------------------------------------------------------------------
    Public Event CountTotals(dblCount As Double, strText As String)

Public Property Let Count(ByVal dblData As Double)
        
    ' ---------------------------------------------------------------------------
    ' used when assigning a value to the property, on the left side of an
    ' assignment.    Syntax: X.Count = 5
    ' ---------------------------------------------------------------------------
    m_dblCount = dblData
        
End Property

Public Property Get Count() As Double
        
    ' ---------------------------------------------------------------------------
    ' used when retrieving value of a property, on the right side of an
    ' assignment.    Syntax: Debug.Print X.Count
    ' ---------------------------------------------------------------------------
    Count = m_dblCount
        
End Property

Public Function IsDriveReady(ByVal strDrive As String) As Boolean
        
    ' ***************************************************************************
    ' Routine:       IsDriveReady
    '
    ' Description:   Test whether a drive is ready.  Do not use a trailing
    '                backslash
    '
    ' Syntax:        IsDriveReady("A:")
    '
    ' Parameters:    strDrive - drive to be tested
    '
    ' Returns:       True or False
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' Capture the first character and convert to uppercase
    ' ---------------------------------------------------------------------------
    strDrive = StrConv(Left$(strDrive, 1), vbUpperCase)
        
    ' ---------------------------------------------------------------------------
    ' Make sure we have the right drive letter syntax
    ' ---------------------------------------------------------------------------
    Select Case Asc(strDrive)
        ' verify this is a letter of the alphabet
        Case 65 To 90
            strDrive = strDrive & ":"
            ' Is this drive available to be accessed
            IsDriveReady = FSO.GetDrive(strDrive).IsReady
                
            ' invalid drive letter passed
        Case Else
            IsDriveReady = False
    End Select
        
    ' ---------------------------------------------------------------------------
    ' Free object from memory
    ' ---------------------------------------------------------------------------
    Set FSO = Nothing
        
End Function

Public Function Create_Dir_Struct(ByVal strPath As String) As Boolean
        
    ' ***************************************************************************
    ' Routine:       Create_Dir_Struct
    '
    ' Description:   Create nested directories.  Must end with a backslash.
    '
    ' Syntax:        Create_Dir_Struct "C:\Program Files\MyDir\Level 1\Level 2\"
    '
    ' Parameters:    strPath = Folder path to be created if it does not exist
    '
    ' Returns:       True or False
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    On Error GoTo Create_Dir_Struct_Errors
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strTmpPath As String
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' See if anything was passed to this routine
    ' ---------------------------------------------------------------------------
    If Len(Trim$(strPath)) = 0 Then
        Create_Dir_Struct = False
        GoTo Normal_Exit
    End If
        
    ' ---------------------------------------------------------------------------
    ' Make sure there is a traling backslash
    ' ---------------------------------------------------------------------------
    strPath = Add_Trailing_Slash(strPath)
    intIndex = 0
        
    ' ---------------------------------------------------------------------------
    ' See if any of the folders have to be created
    ' ---------------------------------------------------------------------------
    Do
        ' get the next path chunk
        intIndex = InStr(intIndex + 1, strPath, "\")
            
        If intIndex > 0 Then
            strTmpPath = Left$(strPath, intIndex - 1)
        Else
            Exit Do
        End If
            
        ' see if this folder exists
        If Not FSO.FolderExists(strTmpPath) Then
            ' Create this folder.
            ' If there is an error, it will be trapped
            ' below and a msgbox displayed.
            FSO.CreateFolder strTmpPath
        End If
    Loop
        
    ' ---------------------------------------------------------------------------
    ' We were successful
    ' ---------------------------------------------------------------------------
    Create_Dir_Struct = True
        
Normal_Exit:
    On Error GoTo 0    ' nullify this error routine
    Set FSO = Nothing
    Exit Function
        
    ' ---------------------------------------------------------------------------
    ' An error occured creating this path
    ' ---------------------------------------------------------------------------
Create_Dir_Struct_Errors:
    TraceOut "ErrNo=[" & CStr(err.Number) & "],ErrDesc=[" & err.Description & _
            "],Error Occured While Trying to Create " & _
            strTmpPath, Error
    Create_Dir_Struct = False
    GoTo Normal_Exit
        
End Function
Public Function Get_Extension(ByVal strFilename As String) As String
        
    ' ---------------------------------------------------------------------------
    ' Capture just the file extension
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
    Get_Extension = FSO.GetExtensionName(strFilename)
    Set FSO = Nothing
        
End Function

Public Function Get_Filename(ByVal strFilename As String) As String
        
    ' ---------------------------------------------------------------------------
    ' Capture just the filename
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
    Get_Filename = FSO.GetFileName(strFilename)
    Set FSO = Nothing
        
End Function

Public Function Get_Path(ByVal strFilename As String) As String
        
    ' ---------------------------------------------------------------------------
    ' Capture full path only
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
    Get_Path = FSO.GetParentFolderName(strFilename)
    Set FSO = Nothing
        
End Function

Public Function Get_Version(ByVal strFilename As String) As String
        
    ' ---------------------------------------------------------------------------
    ' Get file version
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
    Get_Version = FSO.GetFileVersion(strFilename)
    Set FSO = Nothing
        
End Function

Public Function Drive_Exist(ByVal strSearchData As String, _
        ByRef lngDriveType As Long) As Boolean
        
    ' ***************************************************************************
    ' Routine:       Drive_Exist
    '
    ' Description:   Checks to see if the drive we are to query is available
    '
    ' Parameters:    strSearchData - full path\filename
    '
    ' Returns:       True or False
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    On Error GoTo Normal_Exit
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim objDrive As Object
    Dim strHex As String
    Dim strSerial As String
    Dim strDrive As String
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
    lngDriveType = 0
        
    ' ---------------------------------------------------------------------------
    ' separate the drive letter
    ' ---------------------------------------------------------------------------
    If Mid$(strSearchData, 2, 1) = ":" Then
        ' capture just the drive letter and colon
        strDrive = Left$(strSearchData, 2)
    Else
        ' bad input data.  Leave here.
        Drive_Exist = False
        GoTo Normal_Exit
    End If
        
    ' ---------------------------------------------------------------------------
    ' Get the drive information.  If a floppy drive, a disk does not have to be
    ' present to determine its existance.  A disk does have to be present to
    ' gather any additional information.  This is why we use the "IsReady"
    ' property.
    ' ---------------------------------------------------------------------------
    If FSO.DriveExists(strDrive) Then
            
        Set objDrive = FSO.GetDrive(strDrive)
            
        ' See if the drive is ready to be queried
        If IsDriveReady(strDrive) Then
            With DI
                .DriveLetter = objDrive.DriveLetter
                .ShareName = objDrive.ShareName
                .Path = objDrive.Path
                .RootFolder = objDrive.RootFolder
                    
                ' format drive serial number
                strHex = Right$(String$(8, "0") & Hex(objDrive.SerialNumber), 8)
                strSerial = Left$(strHex, 4) & "-" & Right$(strHex, 4)
                .SerialNumber = StrConv(strSerial, vbUpperCase)
                    
                ' get volume label, if any
                .VolumeName = objDrive.VolumeName
                    
                ' Disk space data (handles all sizes)
                .TotalSize = objDrive.TotalSize
                .AvailableSpace = objDrive.AvailableSpace
                .FreeSpace = objDrive.FreeSpace
                    
                ' Drive type:
                '    0  Unknown drive type
                '    1  Local hard drive
                '    2  Floppy or other removable drive
                '    3  Local hard drive (sometimes network drive)
                '    4  Shared Network drive
                '    5  CD-Rom device
                '    6  Virtual memory disk
                .DriveType = objDrive.DriveType
                    
                .FileSystem = objDrive.FileSystem  ' FAT, NTFS, etc.
                .IsReady = IIf(objDrive.IsReady = True, "1", "0")
            End With
                
            lngDriveType = DI.DriveType
            Drive_Exist = True   ' Success
        Else
            Drive_Exist = False  ' not ready to be queried
        End If
    Else
        Drive_Exist = False      ' Drive does not exist
    End If
        
Normal_Exit:
    On Error GoTo 0    ' nullify this error routine
    Set objDrive = Nothing
    Set FSO = Nothing
        
End Function

Public Function File_Exist(strSearchItem As String) As Boolean
        
    ' ***************************************************************************
    ' Routine:       File_Exist
    '
    ' Description:   Test to see if a file exists. Need the fully qualified path.
    '
    ' Parameters:    strSearchItem - Path\filename to be queried.
    '
    ' Returns:       True or False
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim lngType As Long
    Dim strDriveLtr As String
        
    ' ---------------------------------------------------------------------------
    ' See if anything valid was passed here
    ' ---------------------------------------------------------------------------
    If Len(Trim$(strSearchItem)) = 0 Then
        File_Exist = False
        Exit Function
    End If
        
    ' ---------------------------------------------------------------------------
    ' Initialize variables
    ' ---------------------------------------------------------------------------
    lngType = 0
    strDriveLtr = Left$(strSearchItem, 2)
        
    ' ---------------------------------------------------------------------------
    ' see if this is a valid drive letter
    ' ---------------------------------------------------------------------------
    If Not Drive_Exist(strDriveLtr, lngType) Then
        File_Exist = False
        Exit Function
    End If
        
    ' ---------------------------------------------------------------------------
    ' Determine the drive type.
    '    Unknown       = 0  unknown drive type
    '    DRV_FIXED1    = 1  Local hard drive
    '    DRV_REMOVABLE = 2  Floppy or other removeable drive
    '    DRV_FIXED2    = 3  Local hard drive (sometimes network drive)
    '    DRV_NETWORK   = 4  Shared Network drive
    '    DRV_CDROM     = 5  CD-Rom device
    '    DRV_RAMDISK   = 6  Virtual memory disk
    ' ---------------------------------------------------------------------------
    If lngType = 0 Then
        File_Exist = False
        Exit Function
    End If
        
    ' ---------------------------------------------------------------------------
    ' Check to see if this is a valid folder or are we in the root directory?
    ' If the length is equal to three then this must be the root and we have
    ' already tested the drive letter.
    ' ---------------------------------------------------------------------------
    If Len(strSearchItem) > 3 Then
        If Not Folder_Exist(strSearchItem) Then
            File_Exist = False
            Exit Function
        End If
    End If
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' Return either TRUE or FALSE
    ' ---------------------------------------------------------------------------
    If FSO.FileExists(strSearchItem) Then
        File_Exist = True
    Else
        File_Exist = False
    End If
        
    ' ---------------------------------------------------------------------------
    ' Free object from memory
    ' ---------------------------------------------------------------------------
    Set FSO = Nothing
        
End Function

Public Function Folder_Exist(ByVal strFolder As String) As Boolean
        
    ' ***************************************************************************
    ' Routine:       Folder_Exist
    '
    ' Description:   Test for the existance of a folder.
    '
    ' Parameters:    strFolder - full path/folder name
    '
    ' Returns:       True or False
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' See if a file extension is found
    ' ---------------------------------------------------------------------------
    If InStr(1, Right$(strFolder, 5), ".") > 0 Then
        strFolder = FSO.GetParentFolderName(strFolder)
    End If
        
    ' ---------------------------------------------------------------------------
    ' Make sure there are no trailing backslashes
    ' ---------------------------------------------------------------------------
    strFolder = Remove_Trailing_Slash(strFolder)
        
    ' ---------------------------------------------------------------------------
    ' test for the folder existance
    ' ---------------------------------------------------------------------------
    If FSO.FolderExists(strFolder) Then
        Folder_Exist = True
    Else
        Folder_Exist = False
    End If
        
    ' ---------------------------------------------------------------------------
    ' Free object from memory
    ' ---------------------------------------------------------------------------
    Set FSO = Nothing
        
End Function

Public Function Available_Drives() As Variant
        
    ' ***************************************************************************
    ' Routine:       Available_Drives
    '
    ' Description:   Get the list of available drive letters, each separated by
    '                a null character.  (i.e.  a:\ c:\ d:\)
    '
    ' Returns:       Array of available drive letters.  REceiving array must
    '                be a string array.
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 12-MAY-1999  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim strDriveLtrs As String
    Dim arDrives() As String
    Dim intPosition As Integer
        
    ' ---------------------------------------------------------------------------
    ' Use spaces and not null values.  Reduces your chance for errors.
    ' ---------------------------------------------------------------------------
    strDriveLtrs = Space$(MAX_SIZE)
        
    ' ---------------------------------------------------------------------------
    ' capture all the available drives in one long string
    ' ---------------------------------------------------------------------------
    GetLogicalDriveStrings Len(strDriveLtrs), strDriveLtrs
        
    ' ---------------------------------------------------------------------------
    ' get rid of excess trailing blank spaces
    ' ---------------------------------------------------------------------------
    strDriveLtrs = Trim$(strDriveLtrs)
        
    ' ---------------------------------------------------------------------------
    ' get rid of trailing nulls
    ' ---------------------------------------------------------------------------
    intPosition = InStrRev(strDriveLtrs, "\", Len(strDriveLtrs))
    strDriveLtrs = Left$(strDriveLtrs, intPosition)
        
    ' ---------------------------------------------------------------------------
    ' load drive letters into an array
    ' ---------------------------------------------------------------------------
    arDrives = Split(strDriveLtrs, Chr$(0))
        
    ' ---------------------------------------------------------------------------
    ' Return the array of drive letters
    ' ---------------------------------------------------------------------------
    Available_Drives = arDrives()
        
End Function

Public Function IsThisRestricted(strPath As String) As Boolean
        
    ' ***************************************************************************
    ' Routine:       IsThisRestricted
    '
    ' Description:   Determines if we have read, write, delete authority
    '
    ' Parameters:    strPath - area to test
    '
    ' Returns:       True - If this area does not allow updates
    '                False - If this is a valid work area
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 12-MAY-1999  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    On Error GoTo IsThisRestricted_Errors
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim intFile As Integer
    Dim strTmpPath As String
        
    ' ---------------------------------------------------------------------------
    ' initialize local variables
    ' ---------------------------------------------------------------------------
    intFile = FreeFile
    strTmpPath = Trim$(strPath)
        
    ' ---------------------------------------------------------------------------
    ' See if there is a file parameter on the end of these path strings.  If so,
    ' remove it so we can see if we have access to their areas.
    ' ---------------------------------------------------------------------------
    If Len(Trim$(strTmpPath)) > 0 Then
            
        ' add trailing "\" if missing and the dummy filename
        strTmpPath = Add_Trailing_Slash(strTmpPath)
        strTmpPath = strTmpPath & "_Test.$$$"
            
        ' Test to see if we have full authority in this area.
        ' First - Create, Write, Save
        Open strTmpPath For Output As #intFile  ' Open the file
        Print #intFile, "X"                     ' Write to the file
        Close #intFile                          ' Save the file
            
        ' Second - Read, Update
        Open strTmpPath For Append As #intFile  ' Open the file
        Print #intFile, "ABC"                   ' add to the file
        Close #intFile                          ' Save the file
            
        ' Third - Delete
        Kill strTmpPath                         ' Delete the file
        '
        IsThisRestricted = False
    Else
        IsThisRestricted = True
    End If
        
    Exit Function
        
    ' ---------------------------------------------------------------------------
    ' If an error occurs on any of the above steps, we are in a restricted area.
    ' ---------------------------------------------------------------------------
IsThisRestricted_Errors:
    IsThisRestricted = True
    On Error GoTo 0
        
End Function

Public Function Get_FileInfo(strFilename As String) As Boolean
        
    ' ***************************************************************************
    ' Routine:       Get_FileInfo
    '
    ' Description:   Gathers all information about a file
    '
    ' Parameters:    strFileName - fully qualified path\filename to be queried
    '
    ' Returns:       True or False
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    On Error GoTo Normal_Exit
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim objFile As Object
    Dim strAttr As String
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    strAttr = ""
    Get_FileInfo = False
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' Verify something was passed here
    ' ---------------------------------------------------------------------------
    If Len(Trim$(strFilename)) = 0 Then
        Exit Function
    End If
        
    ' ---------------------------------------------------------------------------
    ' get the file data
    ' ---------------------------------------------------------------------------
    Set objFile = FSO.GetFile(strFilename)
        
    ' ---------------------------------------------------------------------------
    ' Store the file information
    ' ---------------------------------------------------------------------------
    With FI
        .Drive = objFile.Drive
        .Path = objFile.Path
        .ParentFolder = objFile.ParentFolder
        .Size = objFile.Size
        .Type = objFile.Type
        .Extension = FSO.GetExtensionName(objFile.Path)
        .DateLastAccessed = objFile.DateLastAccessed
        .DateLastModified = objFile.DateLastModified
        .DateCreated = objFile.DateCreated
        .Attributes = Attribute_String(CLng(objFile.Attributes))
    End With
        
    Get_FileInfo = True
        
Normal_Exit:
    On Error GoTo 0    ' nullify this error routine
    Set objFile = Nothing
    Set FSO = Nothing
        
End Function

Public Function Attribute_String(ByVal lngValue As Long) As String
        
    ' ***************************************************************************
    ' Routine:       Attribute_String
    '
    ' Description:   The numeric value representing the attributes of a
    '                directory or file are passed and the alphabetic
    '                representation is returned.
    '
    '                Syntax::    Attribute_String 35
    '                            Returns   "RHA"
    '
    ' Parameters:    lngValue - numeric attribute representation passed into this
    '                           routine
    '
    ' Returns:       Alphabetic attributes are returned
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim strTmp As String
        
    ' ---------------------------------------------------------------------------
    ' Intialize variables
    ' ---------------------------------------------------------------------------
    strTmp = ""
        
    ' ---------------------------------------------------------------------------
    ' Determine the attribute combination
    ' ---------------------------------------------------------------------------
    Select Case lngValue
        Case 1:  strTmp = "R"
        Case 2:  strTmp = "H"
        Case 3:  strTmp = "RH"
        Case 4:  strTmp = "S"
        Case 5:  strTmp = "RS"
        Case 6:  strTmp = "HS"
        Case 7:  strTmp = "RHS"
        Case 8:  strTmp = "V"
        Case 16: strTmp = "D"
        Case 17: strTmp = "DR"
        Case 18: strTmp = "DH"
        Case 19: strTmp = "DRH"
        Case 20: strTmp = "DS"
        Case 23: strTmp = "DRHS"
        Case 32: strTmp = "A"
        Case 33: strTmp = "RA"
        Case 34: strTmp = "HA"
        Case 35: strTmp = "RHA"
        Case 36: strTmp = "RHSA"
        Case 48: strTmp = "DA"
        Case 49: strTmp = "DRA"
        Case 50: strTmp = "DHA"
        Case 51: strTmp = "DRHA"
        Case 52: strTmp = "DSA"
        Case 53: strTmp = "DRSA"
        Case 54: strTmp = "DHSA"
        Case 55: strTmp = "DRHSA"
        Case 64: strTmp = "Alias"
        Case Else: strTmp = ""
    End Select
        
    ' ---------------------------------------------------------------------------
    ' Return the attribute string
    ' ---------------------------------------------------------------------------
    Attribute_String = strTmp
        
End Function

Public Function Move_File(strSource As String, strTarget As String) As Boolean
        
    On Error GoTo Move_File_Errors
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim intAttr As Integer
    Dim intResp As Integer
    Dim varMsg As Variant
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' See if the target file already exist
    ' ---------------------------------------------------------------------------
    If FSO.FileExists(strTarget) Then
        varMsg = "File already exist at" & vbCrLf & strTarget & vbCrLf
        varMsg = varMsg & FI.DateLastModified & vbCrLf
        varMsg = varMsg & FI.Size & vbCrLf & vbCrLf & "Do you want to replace it?"
            
        SetAttr strTarget, vbNormal ' remove target file attributes
        Kill strTarget              ' delete the target file
            
    End If
        
    ' ---------------------------------------------------------------------------
    ' Move the file
    ' ---------------------------------------------------------------------------
    intAttr = GetAttr(strSource)
    FSO.MoveFile strSource, strTarget
    SetAttr strTarget, intAttr
    Move_File = True
        
Normal_Exit:
    On Error GoTo 0    ' nullify this error routine
    Set FSO = Nothing
    Exit Function
        
Move_File_Errors:
    TraceOut "ErrNo=[" & CStr(err.Number) & "],ErrDesc=[" & err.Description & "],Error Occure While Moving a File " & _
            "FROM=[" & strSource & "],TO=[" & strTarget & "]", Error
    Move_File = False
    GoTo Normal_Exit
        
End Function

Public Function Copy_File(strSource As String, strTarget As String) As Boolean
        
    ' ***************************************************************************
    ' Routine:       Copy_File
    '
    ' Description:   copies a file to another location
    '
    ' Parameters:    strSource - Full qualified path and file name to be copied
    '                strTarget - Full qualified path as a destination
    '
    ' Returns:       TRUE/FALSE
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    On Error GoTo Copy_File_Errors
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim intAttr As Integer
    Dim intResp As Integer
    Dim varMsg As Variant
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' See if the target file already exist
    ' ---------------------------------------------------------------------------
    If FSO.FileExists(strTarget) Then
        varMsg = "File already exist at" & vbCrLf & strTarget & vbCrLf
        varMsg = varMsg & FI.DateLastModified & vbCrLf
        varMsg = varMsg & FI.Size & vbCrLf & vbCrLf & "Do you want to replace it?"
            
            
        SetAttr strTarget, vbNormal ' remove target file attributes
        Kill strTarget              ' delete the target file
            
    End If
        
    ' ---------------------------------------------------------------------------
    ' Copy the file to the new destination
    ' ---------------------------------------------------------------------------
    intAttr = GetAttr(strSource)
    FSO.CopyFile strSource, strTarget, True
    SetAttr strTarget, intAttr
    Copy_File = True
        
        
Normal_Exit:
    On Error GoTo 0    ' nullify this error routine
    Set FSO = Nothing
    Exit Function
        
Copy_File_Errors:
    TraceOut "ErrNO=[" & CStr(err.Number) & "],ErrDesc=[" & err.Description & _
            "],Error Occure While Copying a File " & _
            "FROM=[" & strSource & "],TO=[" & strTarget & "]", Error
        
    Copy_File = False
    GoTo Normal_Exit
        
End Function

Public Function DelTree32(strPath As String) As Boolean
        
    ' ***************************************************************************
    ' Routine:       DelTree32
    '
    ' Description:   Deletes a path and all of its subfolders
    '
    ' Parameters:    strPath - Full path to be queried
    '
    ' Returns:       TRUE/FALSE
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    On Error GoTo DelTree32_Errors
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' Delete complete directory and its sub-directories
    ' ---------------------------------------------------------------------------
    FSO.DeleteFolder strPath, True
    DelTree32 = True
        
        
Normal_Exit:
    On Error GoTo 0    ' nullify this error routine
    Set FSO = Nothing
    Exit Function
        
DelTree32_Errors:
    TraceOut "ErrNo=[" & CStr(err.Number) & "],ErrDesc=[" & err.Description & _
            "],Error Occure While Deleting Path " & strPath, Error
        
    DelTree32 = False
    GoTo Normal_Exit
        
End Function

Public Function Add_Trailing_Slash(ByVal strPath As String) As String
        
    ' ***************************************************************************
    ' Routine:       Add_Trailing_Slash
    '
    ' Description:   Add the trailing backslash from the path if it does not exist
    '
    ' Parameters:    strPath - Full path to be queried
    '
    ' Returns:       Reformatted path
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Initialize variables
    ' ---------------------------------------------------------------------------
    strPath = Trim$(strPath)
        
    ' ---------------------------------------------------------------------------
    ' Test for trailing backslash
    ' ---------------------------------------------------------------------------
    If Right$(strPath, 1) = "\" Then
        Add_Trailing_Slash = strPath
    Else
        Add_Trailing_Slash = strPath & "\"
    End If
        
End Function

Public Function Remove_Trailing_Slash(ByVal strPath As String) As String
        
    ' ***************************************************************************
    ' Routine:       Remove_Trailing_Slash
    '
    ' Description:   Removes the trailing backslash from the path if it exist
    '
    ' Parameters:    strPath - Full path to be queried
    '
    ' Returns:       Reformatted path
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Initialize variables
    ' ---------------------------------------------------------------------------
    strPath = Trim$(strPath)
        
    ' ---------------------------------------------------------------------------
    ' Remove the trailing backslash if it exist
    ' ---------------------------------------------------------------------------
    If Right$(strPath, 1) = "\" Then
        Remove_Trailing_Slash = Left$(strPath, Len(strPath) - 1)
    Else
        Remove_Trailing_Slash = strPath
    End If
        
End Function

Public Function TextSearch(strPath As String, strText As String) As Variant
        
    ' ***************************************************************************
    ' Routine:       TextSearch
    '
    ' Description:   How to search the files in a directory for a piece of
    '                text. The most basic way to do this manually is to loop
    '                through the files in a directory, open each one, and
    '                search the text using a function like InStr.
    '
    ' Parameters:    strPath - path to be searched
    '                strText - Data string to look for
    '
    ' Returns:       An array with the finames that contain this string
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim objStream As Scripting.TextStream
    Dim objFolder As Scripting.Folder
    Dim objFile As Scripting.File
    Dim strTmp As String
    Dim arData() As String
    Dim lngIndex As Long
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' Initialize local variables
    ' ---------------------------------------------------------------------------
    Set objFolder = FSO.GetFolder(strPath)
    lngIndex = 0
    ReDim arData(100)
        
    ' ---------------------------------------------------------------------------
    ' Open each file and look for the text string
    ' ---------------------------------------------------------------------------
    For Each objFile In objFolder.Files
            
        ' open the file for read only
        Set objStream = objFile.OpenAsTextStream(ForReading)
            
        ' dump the contents into a variable
        strTmp = objStream.ReadAll
            
        If InStr(1, strTmp, strText, vbTextCompare) > 0 Then
            arData(lngIndex) = objFile.Path    ' save the file path information
            lngIndex = lngIndex + 1            ' increment the index
                
            ' see if we have to increase the array size
            ' by another 100 elements
            If lngIndex = UBound(arData) Then
                ReDim Preserve arData(lngIndex + 100)
            End If
        End If
    Next
        
    ' ---------------------------------------------------------------------------
    ' See if we found anything. If nothing was found, return "N/A" in the first
    ' element of the array.
    ' ---------------------------------------------------------------------------
    If lngIndex > 0 Then
        lngIndex = lngIndex - 1         ' readjust the index (backup one)
        ReDim Preserve arData(lngIndex) ' resize array to just what was used
    Else
        Erase arData()                  ' empty the array
        ReDim arData(1)                 ' resize to the smallest amount
        arData(0) = "N/A"               ' Show we did not find anything
    End If
        
    ' ---------------------------------------------------------------------------
    ' Return the array of data
    ' ---------------------------------------------------------------------------
    TextSearch = arData()
        
    ' ---------------------------------------------------------------------------
    ' Free objects from memory
    ' ---------------------------------------------------------------------------
    Set objStream = Nothing
    Set objFile = Nothing
    Set objFolder = Nothing
    Set FSO = Nothing
        
End Function

Public Function Get_Totals(ByVal strPath As String, _
        Optional intDataType As Integer = 1, _
        Optional strPattern As String = "*.*", _
        Optional ByVal blnSearchSubfolders As Boolean = True) As String
        
    ' ***************************************************************************
    ' Routine:       Get_Totals
    '
    ' Description:   Gather certain file information as a total
    '
    ' Parameters:    strPath - directory path to be parsed
    '                intDataType - Type of data to be returned
    '                1 - Total amount of data in this path based on a pattern
    '                2 - Total number of files last accessed  on or after a certain date
    '                3 - Total number of files last modified on or after a certain date
    '                4 - Total number of files created on or after a certain date
    '                strPattern - file search pattern (i.e. "*.*" or "*.exe") or
    '                    a date (i.e. "10/02/1995)" or "10/2/95")
    '                blnSearchSubfolders [OPTIONAL] TRUE-[DEFAULT] Search designated
    '                            folder and any subfolders
    '                    FALSE - Search designated folder only, do not search any
    '                            subfolders
    '
    ' Returns:       Either total amount or count
    '
    ' NOTE:  There are certain files that Scrrun.dll will not report on.  One is
    '        the swap file (c:\pagefile.sys").
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim objFolder As Scripting.Folder
    Dim strTmp As String
    Dim strFileList As String
    Dim strList() As String
    Dim strFileData As String
    Dim strInfo As String
    Dim lngIndex As Long
    Dim lngMax As Long
    Dim dblCounter As Double
    Dim intPosition As Integer
        
    ' ---------------------------------------------------------------------------
    ' Initialize variables
    ' ---------------------------------------------------------------------------
    dblCounter = 0
    strInfo = ""
    strFileData = strPattern
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' Make sure something valid was passed here
    ' ---------------------------------------------------------------------------
    If intDataType < 1 Or intDataType > 4 Then
        GoTo Normal_Exit
    ElseIf Len(Trim$(strPath)) = 0 Then
        GoTo Normal_Exit
    End If
        
    ' ---------------------------------------------------------------------------
    ' Make sure a date was passed to this routine, if we are comparing dates.
    ' ---------------------------------------------------------------------------
    If intDataType > 1 Then
        On Error Resume Next
        If Not IsDate(CDate(strFileData)) Then
            GoTo Normal_Exit
        End If
        On Error GoTo 0
    End If
        
    On Error GoTo Get_Totals_Errors
    ' ---------------------------------------------------------------------------
    ' let the private subroutine do all the work
    ' ---------------------------------------------------------------------------
    strFileList = Get_FileList(strPath, blnSearchSubfolders, strPattern)
        
    ' ---------------------------------------------------------------------------
    ' Remove the trailing ";" (semi-colon)
    ' ---------------------------------------------------------------------------
    If Right$(strFileList, 1) = ";" Then
        strFileList = Left$(strFileList, Len(strFileList) - 1)
    End If
        
    ' ---------------------------------------------------------------------------
    ' load the returned data into an array
    ' ---------------------------------------------------------------------------
    strList = Split(strFileList, ";")
    lngMax = UBound(strList)
        
    ' ---------------------------------------------------------------------------
    ' Loop thru the array of file names and collect the file information
    ' ---------------------------------------------------------------------------
    For lngIndex = 0 To lngMax - 1
            
        ' see if there is an empty element is this array
        If Len(Trim$(strList(lngIndex))) = 0 Then
            Exit For   ' leave if we do not have any data
        End If
            
        ' gather the file information
        If Get_FileInfo(strList(lngIndex)) Then
                
            ' evaluate the type of data to be returned
            Select Case intDataType
                             
                Case 1: ' Total amount of data in this path
                    dblCounter = dblCounter + FI.Size
                    RaiseEvent CountTotals(dblCounter, strList(lngIndex))
                        
                Case 2: ' Total number of files accessed on or after a certain date
                        
                    ' capture just the last accessed date
                    intPosition = InStr(1, FI.DateLastAccessed, Chr$(32))
                        
                    If intPosition = 0 Then
                        strTmp = FI.DateLastAccessed
                    Else
                        strTmp = Left$(FI.DateLastAccessed, intPosition - 1)
                    End If
                        
                    ' make the comparison
                    If CDate(strTmp) >= CDate(strFileData) Then
                        dblCounter = dblCounter + 1
                        RaiseEvent CountTotals(dblCounter, strList(lngIndex))
                    End If
                        
                Case 3: ' Total number of files modified on or after a certain date
                        
                    ' capture just the last modified date
                    intPosition = InStr(1, FI.DateLastModified, Chr$(32))
                        
                    If intPosition = 0 Then
                        strTmp = FI.DateLastModified
                    Else
                        strTmp = Left$(FI.DateLastModified, intPosition - 1)
                    End If
                        
                    ' make the comparison
                    If CDate(strTmp) >= CDate(strFileData) Then
                        dblCounter = dblCounter + 1
                        RaiseEvent CountTotals(dblCounter, strList(lngIndex))
                    End If
                        
                Case 4: ' Total number of files created on or after a certain date
                        
                    ' capture just the created date
                    intPosition = InStr(1, FI.DateCreated, Chr$(32))
                        
                    If intPosition = 0 Then
                        strTmp = FI.DateCreated
                    Else
                        strTmp = Left$(FI.DateCreated, intPosition - 1)
                    End If
                        
                    ' make the comparison
                    If CDate(strTmp) >= CDate(strFileData) Then
                        dblCounter = dblCounter + 1
                        RaiseEvent CountTotals(dblCounter, strList(lngIndex))
                    End If
                        
                Case 5: ' Total number of a particular type of file.
                    ' A text compare is performed because it does not
                    ' distinguish between upper and lower case characters
                    If StrComp(FI.Type, strFileData, vbTextCompare) = 0 Then
                        dblCounter = dblCounter + 1
                        RaiseEvent CountTotals(dblCounter, strList(lngIndex))
                    End If
            End Select
        End If
    Next
        
Normal_Exit:
    ' ---------------------------------------------------------------------------
    ' Convert to a string before leaving.
    ' ---------------------------------------------------------------------------
    On Error GoTo 0    ' nullify this error routine
    Get_Totals = CStr(dblCounter)
        
    ' ---------------------------------------------------------------------------
    ' Free objects from memory
    ' ---------------------------------------------------------------------------
    Set objFolder = Nothing
    Set FSO = Nothing
    Exit Function
        
Get_Totals_Errors:
    TraceOut "ErrNo=[" & CStr(err.Number) & "],ErrDesc=[" & err.Description & _
            "],File=[" & strList(lngIndex) & _
            "],Size=[" & Format$(FI.Size, "#,0") & _
            "],Created=[" & FI.DateCreated & _
            "],Modified=[" & FI.DateLastModified & _
            "],Accessed=[" & FI.DateLastAccessed, Error
        
    dblCounter = 0
    GoTo Normal_Exit
        
End Function

Public Function Get_FileList(ByVal strFolder As String, _
        Optional ByVal blnSearchSubfolders As Boolean = True, _
        Optional strPattern As String = "*.*") As Variant
        
    ' ***************************************************************************
    ' Routine:       Get_FileList
    '
    ' Description:   procedure that returns a file list as a comma-delimited
    '                list of files
    '
    ' Parameters:    strFolder - Full path to folder
    '                blnSearchSubfolders [OPTIONAL] TRUE-[DEFAULT] Search designated
    '                          folder and any subfolders
    '                          FALSE - Search designated folder only, do not
    '                          search any subfolders
    '                strPattern - [OPTIONAL] [DEFAULT] "*.*" (all files)
    '
    ' Returns:       A string of files delimited by ";"
    '
    ' Syntax:        Get_FileList("C:\Zip", True, "*.zip")
    '                Ex:  Return a list of all ZIP files
    '                C:\Zip\File1.zip;C:\Zip\File2.zip;...C:\Zip\Level9\File1.zip;
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' 26-MAR-2001  Kenneth Ives  kenaso@home.com
    '              Modified search pattern capabilities and documented
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim objFile As Scripting.File
    Dim objFolder As Scripting.Folder
    Dim objSubFolder As Scripting.Folder
    Dim strFileExt As String
    Dim strTestExt As String
    Dim strNewFolder As String
    Dim blnGetAllFiles As Boolean
        
    ' ---------------------------------------------------------------------------
    ' see if the user cancelled processing
    ' ---------------------------------------------------------------------------
    DoEvents
    If CancelProcessing Then
        Exit Function
    End If
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' get the starting folder
    ' ---------------------------------------------------------------------------
    Set objFolder = FSO.GetFolder(strFolder)
        
    ' ---------------------------------------------------------------------------
    ' Test the pattern we are looking for
    ' ---------------------------------------------------------------------------
    If strPattern = "*.*" Then
        blnGetAllFiles = True
    ElseIf InStr(strPattern, "*") = 0 Then
        blnGetAllFiles = True
    Else
        ' looking for a specific pattern
        blnGetAllFiles = False
        strTestExt = StrConv(FSO.GetExtensionName(strPattern), vbLowerCase)
    End If
        
    ' ---------------------------------------------------------------------------
    ' Check all the files in this directory
    ' ---------------------------------------------------------------------------
    For Each objFile In objFolder.Files
            
        ' see if the user cancelled processing
        DoEvents
        If CancelProcessing Then
            Exit For
        End If
            
        If blnGetAllFiles Then
            ' capture all files
            Get_FileList = Get_FileList & objFile.Path & ";"
        Else
            ' look for a specific extension
            strFileExt = StrConv(FSO.GetExtensionName(objFile.Path), vbLowerCase)
                
            ' See if the file extensions match
            If StrComp(strTestExt, strFileExt, vbTextCompare) = 0 Then
                Get_FileList = Get_FileList & objFile.Path & ";"
                m_dblCount = m_dblCount + 1
                RaiseEvent CountFiles(m_dblCount, objFile.Path)
            End If
        End If
    Next
        
    ' ---------------------------------------------------------------------------
    ' see if the user cancelled processing
    ' ---------------------------------------------------------------------------
    DoEvents
    If CancelProcessing Then
        Get_FileList = ""
    Else
        ' if requested, also search subdirectories.
        If blnSearchSubfolders Then
                
            For Each objSubFolder In objFolder.SubFolders
                    
                ' see if the user cancelled processing
                DoEvents
                If CancelProcessing Then
                    Exit For
                End If
                    
                strNewFolder = objSubFolder
                    
                ' Do recursive calls from here.
                Get_FileList = Get_FileList & _
                        Get_FileList(strNewFolder, blnSearchSubfolders, strPattern)
            Next
        End If
    End If
        
Normal_Exit:
    ' ---------------------------------------------------------------------------
    ' Free objects from memory
    ' ---------------------------------------------------------------------------
    Set objFile = Nothing
    Set objFolder = Nothing
    Set objSubFolder = Nothing
    Set FSO = Nothing
        
End Function

Public Function Get_DirList(ByVal strPath As String) As String
        
    ' ***************************************************************************
    ' Routine:       Get_DirList
    '
    ' Description:   procedure that returns a list of folders in a string
    '                delimited by a semi-colon
    '
    ' Parameters:    strPath - path to be parsed
    '
    ' Returns:       A directory list
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-NOV-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim objFolder As Scripting.Folder
    Dim objSubFolder As Scripting.Folder
    Dim strList As String
        
    ' ---------------------------------------------------------------------------
    ' Initialize variables
    ' ---------------------------------------------------------------------------
    strList = ""
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
    Set objFolder = FSO.GetFolder(strPath)
        
    ' ---------------------------------------------------------------------------
    ' list all the folders within this folder
    ' ---------------------------------------------------------------------------
    For Each objSubFolder In objFolder.SubFolders
        ' see if the user cancelled processing
        DoEvents
        If CancelProcessing Then
            Exit For
        End If
            
        strList = strList & objSubFolder.Path & ";"
    Next
        
    ' ---------------------------------------------------------------------------
    ' see if the user cancelled processing
    ' ---------------------------------------------------------------------------
    DoEvents
    If CancelProcessing Then
        strList = ""
    Else
        ' Remove the last ";" (semi-colon)
        If Right$(strList, 1) = ";" Then
            strList = Left$(strList, Len(strList) - 1)
        End If
    End If
        
    ' ---------------------------------------------------------------------------
    ' Return the list of directories in a semi-colon delimited string
    ' ---------------------------------------------------------------------------
    Get_DirList = strList
        
    ' ---------------------------------------------------------------------------
    ' Free objects from memory
    ' ---------------------------------------------------------------------------
    Set objSubFolder = Nothing
    Set objFolder = Nothing
    Set FSO = Nothing
        
End Function

Public Function Create_Tmp_Filename(Optional strTmpFolder As String = vbNullString) As String
        
    ' ***************************************************************************
    ' Routine:       Create_Tmp_Filename
    '
    ' Description:   System generated temporary folder and file.  The folder
    '                will be located in the Windows default temp directory and
    '                is system generated.
    '
    '                Found this example at http://www.vbapi.com/
    '
    ' Parameters:    strTmpFolder - Path to the temporary folder.
    '
    ' Returns:       Unique name of a temporary file
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 02-DEC-1999  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim strPrefix As String
    Dim strTmpFile As String
    Dim lngRetCode As Long
        
    ' ---------------------------------------------------------------------------
    ' Use spaces and not null values.  Reduces your errors.
    ' ---------------------------------------------------------------------------
    strPrefix = ""
    strTmpFile = Space$(MAX_SIZE)
        
    ' ---------------------------------------------------------------------------
    ' See if the temp folder path was passed.  If not, go get it.
    ' ---------------------------------------------------------------------------
    If Len(Trim$(strTmpFolder)) = 0 Then
        strTmpFolder = Space$(MAX_SIZE)                  ' initialize buffer
        lngRetCode = GetTempPath(MAX_SIZE, strTmpFolder) ' read the path name
            
        ' extract data from the variable (ex:  "C:\Temp\")
        If lngRetCode Then
            ' We found the Windows default Temp folder.  Remove
            ' the trialing nulls and append the name of our
            ' temporary folder.
            strTmpFolder = Remove_Nulls(strTmpFolder)
            strTmpFolder = Add_Trailing_Slash(strTmpFolder)
        Else
            ' use root directory of drive C:\
            strTmpFolder = "C:\"
        End If
    End If
        
    ' ---------------------------------------------------------------------------
    ' Create a unique temporary file name.  I define the prefix.  A four
    ' digit hex number is returned.  (Ex: 49489 becomes C151)
    ' ---------------------------------------------------------------------------
    strPrefix = Create_Tmp_Name(4)              ' Ex:  "" will return ""
    lngRetCode = GetTempFileName(strTmpFolder, strPrefix, 0&, strTmpFile)
        
    ' ---------------------------------------------------------------------------
    ' Note that the file is also created for you. Extract data from the variable
    ' (Ex:  "C:\Temp\_resC151.TMP")
    ' ---------------------------------------------------------------------------
    strTmpFile = Remove_Nulls(strTmpFile)
        
    ' ---------------------------------------------------------------------------
    ' Return the path and name of the temp file
    ' ---------------------------------------------------------------------------
    Create_Tmp_Filename = strTmpFile
        
End Function

Public Function Create_Tmp_Folder() As String
        
    ' ***************************************************************************
    ' Routine:       Create_Tmp_Folder
    '
    ' Description:   System generated temporary folder and file.  The folder
    '                will be located in the Windows default temp directory and
    '                is system generated.
    '
    '                Found this example at http://www.vbapi.com/
    '
    ' Returns:       Name of the temp folder
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 02-DEC-1999  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim strUniqueName As String
    Dim strTmpFolder As String
    Dim lngRetCode As Long
        
    ' ---------------------------------------------------------------------------
    ' Instantsiate Scripting object
    ' ---------------------------------------------------------------------------
    Set FSO = New Scripting.FileSystemObject
        
    ' ---------------------------------------------------------------------------
    ' See if we can find the Windows temporary directory path.
    ' ---------------------------------------------------------------------------
    strUniqueName = Create_Tmp_Name(8)               ' Create a unique folder name
    strTmpFolder = Space$(MAX_SIZE)                  ' initialize buffer
    lngRetCode = GetTempPath(MAX_SIZE, strTmpFolder) ' read the path name
        
    ' ---------------------------------------------------------------------------
    ' extract data from the variable (ex:  "C:\Temp\")
    ' ---------------------------------------------------------------------------
    If lngRetCode Then
        ' We found the Windows default Temp folder.  Remove
        ' the trialing nulls and append the name of our
        ' temporary folder.
        strTmpFolder = Remove_Nulls(strTmpFolder)
        strTmpFolder = Add_Trailing_Slash(strTmpFolder)
        strTmpFolder = strTmpFolder & strUniqueName
    Else
        ' We did not find the Windows default temp folder
        strTmpFolder = "C:\" & strUniqueName
    End If
        
    ' ---------------------------------------------------------------------------
    ' Crate the new temp subfolder
    ' ---------------------------------------------------------------------------
    strTmpFolder = Remove_Trailing_Slash(strTmpFolder)
        
    If Not FSO.FolderExists(strTmpFolder) Then
        ' create the temp folder within the Windows temp folder
        FSO.CreateFolder (strTmpFolder)
    End If
        
    ' ---------------------------------------------------------------------------
    ' Return the full path to the new temp folder
    ' ---------------------------------------------------------------------------
    strTmpFolder = Add_Trailing_Slash(strTmpFolder)
    Create_Tmp_Folder = strTmpFolder
    Set FSO = Nothing
        
End Function

Public Function Remove_Nulls(strInput As String) As String
        
    ' ***************************************************************************
    ' Routine:       Remove_Nulls
    '
    ' Description:   Receives a data string and looks for the first null.  If
    '                found, the string is truncated at that point and returned.
    '
    ' Parameters:    strInput - Input string to be inspected
    '
    ' Returns:       Data string with the trailing nulls
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 23-FEB-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim intPosition As Integer
        
    ' ---------------------------------------------------------------------------
    ' Look for first null in the data string
    ' ---------------------------------------------------------------------------
    intPosition = InStr(1, strInput, Chr$(0))
        
    ' ---------------------------------------------------------------------------
    ' if we found one, then save all the data up to that position
    ' ---------------------------------------------------------------------------
    If intPosition Then
        Remove_Nulls = Left$(strInput, intPosition - 1)
    Else
        Remove_Nulls = strInput     ' no nulls found
    End If
        
End Function

Public Function Create_Tmp_Name(intLength As Integer) As String
        
    ' ***************************************************************************
    ' Routine:       Create_Tmp_Name
    '
    ' Description:   System generated temporary name.  User determines the length
    '                of the name, not to exceed eight characters.
    '
    ' Parameters:    intLength - length of character string desired
    '
    ' Returns:       New name
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 09-APR-2000  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim strTmp As String
    Dim intChar As Integer
    Dim intIndex As Integer
        
    ' ---------------------------------------------------------------------------
    ' Initialize variables
    ' ---------------------------------------------------------------------------
    strTmp = "_"  ' underscore is the first character
        
    ' ---------------------------------------------------------------------------
    ' Test length passed.  0=exit function;  1=save what we got and leave
    ' ---------------------------------------------------------------------------
    If intLength < Len(strTmp) Then
        Exit Function
    ElseIf intLength = Len(strTmp) Then
        Create_Tmp_Name = strTmp
        Exit Function
    End If
        
    ' ---------------------------------------------------------------------------
    ' Create string using 0-9, A-Z, a-z characters only
    ' ---------------------------------------------------------------------------
    Do
        intChar = Int(Rnd2(48, 122))      ' create character 0-9, A-Z, a-z
            
        Select Case intChar               ' test character creatation
            Case 58 To 64, 91 To 96
                intChar = intChar + 7 ' add 7 to unwanted char value
        End Select
            
        strTmp = strTmp & Chr$(intChar)   ' append new char to output string
            
    Loop Until Len(strTmp) = intLength
        
    ' ---------------------------------------------------------------------------
    ' Return the new name
    ' ---------------------------------------------------------------------------
    Create_Tmp_Name = StrConv(strTmp, vbLowerCase)
        
End Function

Public Function Rnd2(sngLow As Single, sngHigh As Single) As Single
        
    ' ***************************************************************************
    ' Routine:       Rnd2
    '
    ' Description:   Create a random value between two values.  We are assuming
    '                that the random number generator has already been seeded.
    '
    ' Parameters:    sngLow  - Low end value
    '                sngHign - High end value
    '
    ' Return Values: A random generated value
    '
    ' ===========================================================================
    '    DATE      NAME / eMAIL
    '              DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 26-JUN-1999  The VB2TheMax Team  fbalena@vb2themax.com
    '              Routine created
    ' 03-OCT-2000  Kenneth Ives  kenaso@home.com
    '              Modified and documented
    ' ***************************************************************************
    DoEvents        ' gives the system clock time to advance a few milliseconds
    Randomize (CDbl(Now()) + Timer)     ' Reseed the VB random number generator
    Rnd2 = (Rnd * (sngHigh - sngLow)) + sngLow
        
End Function

Public Function BrowseForFolder(Optional frm As Form = Nothing, _
        Optional strPrompt As String = "") As String
        
    ' ***************************************************************************
    ' Routine:       BrowseForFolder
    '
    ' Description:   Opens the Browse Dialog box so a folder may be slected by
    '                the user.  Both parameters are optional.
    '
    ' Parameters:    frm - name of form to align the upper left corner of the
    '                      browse dialog box
    '                strPrompt - Additional title line in the Browse dialog box
    '
    ' Returns:       Full path to the folder
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 23-FEB-1999  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim lngFolderList As Long
    Dim lngRetCode As Long
    Dim strTmpPath As String
    Dim BI As BrowseInfo
        
    ' ---------------------------------------------------------------------------
    ' Initialize variables
    ' ---------------------------------------------------------------------------
    strTmpPath = ""
    lngFolderList = 0
        
    ' ---------------------------------------------------------------------------
    ' Set up Browse Dialog Box parameters
    ' ---------------------------------------------------------------------------
    With BI
        ' See if form name was passed.  if not use a 0
        If frm Is Nothing Then
            .hWndOwner = 0&
        Else
            .hWndOwner = frm.hwnd
        End If
            
        ' strPrompt will be the title on the Browse dialog box
        .lpszTitle = lstrcat(strPrompt, "")
            
        ' Display only directories
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
        
    ' ---------------------------------------------------------------------------
    ' display the browse dialog box and  create an index of the folders in the
    ' dialog box
    ' ---------------------------------------------------------------------------
    lngFolderList = SHBrowseForFolder(BI)
        
    ' ---------------------------------------------------------------------------
    ' If a folder was highlighted then format the folder name so it can be
    ' returned
    ' ---------------------------------------------------------------------------
    If lngFolderList Then
            
        ' set up a pre-padded buffer area for the folder name.
        ' Use spaces and not null values.  Reduces your errors.
        strTmpPath = Space$(MAX_SIZE)
            
        ' Get the name of the folder from the list
        lngRetCode = SHGetPathFromIDList(lngFolderList, strTmpPath)
        Call CoTaskMemFree(lngFolderList)
            
        ' Strip any null characters from the folder name
        strTmpPath = Remove_Nulls(strTmpPath)
    End If
        
    ' ---------------------------------------------------------------------------
    ' return the formatted folder name
    ' ---------------------------------------------------------------------------
    BrowseForFolder = strTmpPath
        
End Function

Public Function Shrink_2_Fit(ByVal strPath_to_Resize As String, _
        intMaxLen As Integer) As String
        
    ' ***************************************************************************
    ' Routine:       Shrink_2_Fit
    '
    ' Description:   This function will shorten a directory name to the length
    '                passed to the Max parameter.
    '
    ' Syntax:
    '        Shrink_2_Fit "C:\Program Files\Navigator\Programs\Bookmark.htm", 30
    '        Returns -->  "C:\...\Programs\Bookmark.htm"
    '
    ' Parameters:    strPath_to_Resize - Path to be resized for display
    '                intMaxLen - Maximum length of the return string
    '
    ' Returns:       Resized path
    '
    ' ===========================================================================
    '    DATE      NAME / DESCRIPTION
    ' -----------  --------------------------------------------------------------
    ' 23-FEB-1998  Kenneth Ives  kenaso@home.com
    '              Routine created
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim intPosition As Integer
    Dim intStart As Integer
    Dim intNeeded As Integer
    Dim intLength As Integer
    Dim strTmp As String
    Dim strPrefix As String
        
    ' ---------------------------------------------------------------------------
    ' Initialize variables
    ' ---------------------------------------------------------------------------
    strTmp = Trim$(strPath_to_Resize)
        
    ' ---------------------------------------------------------------------------
    ' if the string is equal to or less than the desired length then leave
    ' this routine
    ' ---------------------------------------------------------------------------
    If Len(strTmp) <= intMaxLen Then
        Shrink_2_Fit = strTmp
        Exit Function
    End If
        
    ' ---------------------------------------------------------------------------
    ' Prepare the drive and the elipces.  Start in the third position in case
    ' this is a UNC (Universal Naming Convention) path
    ' ---------------------------------------------------------------------------
    If Left$(strTmp, 2) = "\\" Or Mid$(strTmp, 2, 1) = ":" Then
        intPosition = InStr(3, strTmp, "\")            ' search for next backslash
        intStart = intPosition + 1                     ' calc new starting position
        intNeeded = intMaxLen - (intPosition + 3)      ' Add for 3 elipces
    Else
        Shrink_2_Fit = strTmp   ' do not know what this is.  get rid of it.
        Exit Function
    End If
        
    ' ---------------------------------------------------------------------------
    ' Build the prefix.  A drive letter or server_name, backslash and 3 elipces
    ' Example:   "C:\..." or "\\Server\..."
    ' ---------------------------------------------------------------------------
    strPrefix = Left$(strTmp, intPosition) & "..."
        
    ' ---------------------------------------------------------------------------
    ' prepare the suffix part of the path to the desired length
    ' ---------------------------------------------------------------------------
    Do
        intPosition = InStr(intStart, strTmp, "\") ' look for a backslash
            
        If intPosition > 0 Then
            strTmp = Mid$(strTmp, intPosition)      ' Reformat string length
            intLength = Len(strTmp)
        Else
            Exit Do                                ' could not find "", so leave
        End If
            
    Loop Until intNeeded >= intLength
        
    ' ---------------------------------------------------------------------------
    ' Return the readjusted data string
    ' ---------------------------------------------------------------------------
    Shrink_2_Fit = strPrefix & strTmp
        
End Function

Public Function Get_Disk_Space(ByVal strPath As String, _
        dblTotalSpace As Double, _
        dblFreeSpace As Double, _
        dblUsedSpace As Double) As Boolean
        
    ' ***************************************************************************
    ' Routine:       Get_Disk_Space
    '
    ' Description:   Get the total number of bytes, total number of free bytes,
    '                and total number of used bytes on a hard drive.  This  will
    '                also determine which API call to make based on the size of
    '                the drive.
    '
    ' Parameters:    strDriveLtr   - drive to be queried (ex:  C:)
    '                dblTotalSpace - Value to be returned
    '                dblFreeSpace  - Value to be returned
    '                dblUsedSpace  - Value to be returned
    '
    ' Returns:       Total number of bytes, total number of free bytes, and total
    '                number of used bytes on a drive.
    '
    ' ===========================================================================
    '    DATE      NAME             DESCRIPTION
    ' -----------  ---------------  ---------------------------------------------
    ' 03-MAR-2000  Kenneth Ives     Wrote routine
    ' 14-MAR-2000  Kenneth Ives     Modified and documented to handle under and
    '                               over 2gb drives
    ' ***************************************************************************
        
    ' ---------------------------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------------------------
    Dim lngHandle As Long
    Dim lngRetCode As Long
    Dim lngProcAddress As Long
    Dim lngSectorsPerCluster As Long
    Dim lngBytesPerSector As Long
    Dim lngNumberOfFreeClusters As Long
    Dim lngTotalNumberOfClusters As Long
    Dim curFreeBytesToCaller As Currency
    Dim curTotalFreeBytes As Currency
    Dim curTotalBytes As Currency
    Dim strDriveLtr As String
        
    '  Const TWO_GB                  As Double = 2147483647#
    Const TWO_GB As Double = 2200000000#
        
    ' ---------------------------------------------------------------------------
    ' Initialize variables
    ' ---------------------------------------------------------------------------
    dblFreeSpace = 0
    dblTotalSpace = 0
    dblUsedSpace = 0
    curFreeBytesToCaller = 0
    curTotalFreeBytes = 0
    curTotalBytes = 0
    lngSectorsPerCluster = 0
    lngBytesPerSector = 0
    lngNumberOfFreeClusters = 0
    lngTotalNumberOfClusters = 0
    strDriveLtr = Trim$(strPath)
        
    ' ---------------------------------------------------------------------------
    ' Verify data was passed
    ' ---------------------------------------------------------------------------
    If Len(Trim$(strDriveLtr)) = 0 Then
        Exit Function
    End If
        
    ' ---------------------------------------------------------------------------
    ' Remove any trailing backslashes
    ' ---------------------------------------------------------------------------
    strDriveLtr = Left$(strDriveLtr, 1) & ":"
        
    On Error GoTo Over_2gb_Not_Supported
    ' ---------------------------------------------------------------------------
    ' First, determine if we can call the GetDiskFreeSpaceEx function.
    ' ---------------------------------------------------------------------------
    lngHandle = LoadLibrary("kernel32.Dll")
        
    ' ---------------------------------------------------------------------------
    ' if the return code is not zero, we were successful
    ' ---------------------------------------------------------------------------
    If lngHandle Then
        lngProcAddress = GetProcAddress(lngHandle, "GetDiskFreeSpaceExA")
            
        If lngProcAddress Then
                
            ' Make API call to query the hard drive for space information.
            ' if there is an error, we will assume that this operating
            ' system does not support hard drives over 2gb and then jump
            ' to the area where we can make the correct API call.
            lngRetCode = GetDiskFreeSpaceEx(strDriveLtr, _
                    curFreeBytesToCaller, _
                    curTotalBytes, _
                    curTotalFreeBytes)
                
            ' Just to be safe, decrement Dll's usage counter.
            FreeLibrary lngHandle
                
            ' If we were successful, then multiply the results
            ' by 10000 to move the decimal to the right 4 positions
            If lngRetCode Then
                ' Calculate size of the hard drive
                dblTotalSpace = CDbl(curTotalBytes * 10000)
                ' Calculate the amount of free space
                dblFreeSpace = CDbl(curTotalFreeBytes * 10000)
                ' Calculate the amount of used space
                dblUsedSpace = dblTotalSpace - dblFreeSpace
            End If
        End If
            
        ' Just to be safe, decrement Dll's usage counter.
        FreeLibrary lngHandle
    End If
        
    ' ---------------------------------------------------------------------------
    ' If an error brought us to here then make sure we reset all variables
    ' ---------------------------------------------------------------------------
Over_2gb_Not_Supported:
    If err.Number <> 0 Then
        err.Clear           ' reset error code
        dblTotalSpace = 0   ' empty variables
        dblFreeSpace = 0
        dblUsedSpace = 0
        lngRetCode = 0
    End If
        
    On Error GoTo 0
    On Error GoTo Get_Disk_Space_Errors
    ' ---------------------------------------------------------------------------
    ' If this drive is less than 2GB in total size then reset values and make
    ' the correct API call.
    ' ---------------------------------------------------------------------------
    If dblTotalSpace <= TWO_GB Then
        dblTotalSpace = 0
        dblFreeSpace = 0
        dblUsedSpace = 0
        lngRetCode = 0
    Else
        Get_Disk_Space = True
        Exit Function
    End If
        
    ' ---------------------------------------------------------------------------
    ' if we get here, GetDiskFreeSpaceEx isn't available or raised an error or
    ' the drive is less than 2GB in size.
    ' ---------------------------------------------------------------------------
    lngRetCode = GetDiskFreeSpace(strDriveLtr, lngSectorsPerCluster, _
            lngBytesPerSector, lngNumberOfFreeClusters, _
            lngTotalNumberOfClusters)
        
    ' ---------------------------------------------------------------------------
    ' Test the return code to see if we had a successful call
    ' ---------------------------------------------------------------------------
    If lngRetCode Then
        ' Calculate size of the hard drive
        dblFreeSpace = CDbl(lngNumberOfFreeClusters * lngSectorsPerCluster * _
                lngBytesPerSector)
        ' Calculate the amount of free space
        dblTotalSpace = CDbl(lngTotalNumberOfClusters * lngSectorsPerCluster * _
                lngBytesPerSector)
        ' Calculate the amount of used space
        dblUsedSpace = dblTotalSpace - dblFreeSpace
        Get_Disk_Space = True
    Else
        Get_Disk_Space = False
    End If
        
Normal_Exit:
    Exit Function
        
Get_Disk_Space_Errors:
    Get_Disk_Space = False
    Resume Normal_Exit
        
End Function

Private Sub Class_Initialize()
        
    ' ---------------------------------------------------------------------------
    ' Initialize holding areas
    ' ---------------------------------------------------------------------------
    Dim FI_dummy As FileInfo
    Dim DI_dummy As DriveInfo
        
    FI = FI_dummy      ' fill structures with empty values
    DI = DI_dummy
    m_dblCount = 0
        
End Sub

Private Sub Class_Terminate()
        
    ' ---------------------------------------------------------------------------
    ' Initialize holding areas
    ' ---------------------------------------------------------------------------
    Dim FI_dummy As FileInfo
    Dim DI_dummy As DriveInfo
        
    FI = FI_dummy      ' fill structures with empty values
    DI = DI_dummy
    m_dblCount = 0
        
End Sub

