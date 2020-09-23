Attribute VB_Name = "Module1"
Private Const MOVEFILE_REPLACE_EXISTING = &H1
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const CREATE_NEW = 1
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' Following code executes a file with a known/registered extension
Private Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' Used to wordwrap RTF boxes & find listbox/combobox items
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Const WM_USER = &H400
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)

Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_FINDSTRINGEXACT As Long = &H158
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
Private Const LB_FINDSTRING As Long = &H18F

Public lReadSize As Long

'-----------------------------------------------------------
' FUNCTION: GetTempPath
'
' Calls the windows API to get the windows GetTempPath  and
' ensures that a trailing dir separator is present
'
' Returns: The windows directory
'-----------------------------------------------------------
'
Public Function GetTempFile(bToDelete As Boolean) As String
    Dim strBuf As String, nPos As Long, sTempFolder As String

    strBuf = Space$(256)
    '
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetTempPath(256, strBuf) Then
      nPos = InStr(strBuf, vbNullChar)
      If nPos > 0 Then
          sTempFolder = Left$(strBuf, nPos - 1)
      Else
          sTempFolder = "C:\"
      End If

        If Right$(sTempFolder, 1) <> "\" Then sTempFolder = sTempFolder & "\"
    End If
    GetTempFile = Space$(256)
    If bToDelete Then strBuf = "~Lv" Else strBuf = "~QT"
    GetTempFileName sTempFolder, strBuf, 0, GetTempFile
    GetTempFile = Left$(GetTempFile, InStr(1, GetTempFile, Chr$(0)) - 1)
    'Set the file attributes
End Function

Public Function RetrieveFileHandle(sFileName As String) As Long
RetrieveFileHandle = CreateFile(sFileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
End Function

Public Function ReadFilePart(sSourceFile As String, sDestFileName As String, _
    iStartPos As Long, iStopPos As Long, iEndPos As Long, iMode As Integer, _
    iDirection As Integer, iAnchor As Integer) As Long
    
    Dim hOrgFile As Long, hNewFile As Long, bBytes() As Byte
    Dim nSize As Long, lRet As Long, lBytesRead As Long
    Dim lPointer As Long, lEndPoint As Long, lPortion As Long, lFileSize As Long
    
    On Error GoTo BadRead
    'Open the files
    hNewFile = CreateFile(sDestFileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, 0, 0)
    hOrgFile = RetrieveFileHandle(sSourceFile)

    'Get the file size
    lFileSize = GetFileSize(hOrgFile, 0)
    If iMode = 0 Then   ' reading by percent
        lPortion = lFileSize * (iEndPos / 100)
    Else
        lPortion = iEndPos
    End If
    If lPortion > 1024# * 10000# Then
        lPortion = 1024# * 10000#
        MsgBox "Note. Maximum read size is 10 mb of data per view.", vbInformation + vbOKOnly
    End If
    Select Case iAnchor
    Case 0: ' beginning of file
        lPointer = 0
        lEndPoint = lPortion
    Case 1: ' current position
        lPointer = iStartPos
        lEndPoint = lPortion
    Case 2: ' end of file
        lPointer = lFileSize - lPortion
        lEndPoint = lFileSize
    End Select
    If iDirection = 1 Then lPointer = lPointer - lPortion
    If lPointer <= 0 Then lPointer = 0
    If lEndPoint + lPointer > lFileSize Then lEndPoint = lFileSize - lPointer
    If lFileSize - (lEndPoint + lPointer) < 1024 Then lEndPoint = lFileSize - lPointer
    iStartPos = lPointer
    iStopPos = lPointer + lEndPoint
    If iStopPos = lFileSize Then iStopPos = -1
    
    SetFilePointer hOrgFile, lPointer, 0, FILE_BEGIN        'Set the file pointer
    ReDim bBytes(1 To lReadSize) As Byte       'Create an array of bytes
    Do While nSize < lEndPoint
        If nSize + lReadSize > lEndPoint Then ReDim bBytes(1 To lEndPoint - nSize)
        ReadFile hOrgFile, bBytes(1), UBound(bBytes), lBytesRead, ByVal 0&  'Read from the file
        If lBytesRead = 0 Then MsgBox "Error reading file ...": Exit Do     'Check for errors
        WriteFile hNewFile, bBytes(1), lBytesRead, lRet, ByVal 0&           'Write to the file
        If lRet <> lBytesRead Then MsgBox "Error writing file ...": Exit Do 'Check for errors
        nSize = nSize + lBytesRead
    Loop
    'Close the files
CloseHandles:
    If hOrgFile Then CloseHandle hOrgFile
    If hNewFile Then CloseHandle hNewFile
Exit Function

BadRead:
MsgBox Err.Description, vbExclamation + vbOKOnly
Resume CloseHandles
End Function

Public Sub ShellSortArray(vArray As Variant)
  Dim lLoop1 As Long
  Dim lHold As Long
  Dim lHValue As Long
  Dim lTemp As Variant

  lHValue = LBound(vArray)
  Do
    lHValue = 3 * lHValue + 1
  Loop Until lHValue > UBound(vArray)
  Do
    lHValue = lHValue / 3
    For lLoop1 = lHValue + LBound(vArray) To UBound(vArray)
      lTemp = vArray(lLoop1)
      lHold = lLoop1
      Do While vArray(lHold - lHValue) > lTemp
        vArray(lHold) = vArray(lHold - lHValue)
        lHold = lHold - lHValue
        If lHold < lHValue Then Exit Do
      Loop
      vArray(lHold) = lTemp
    Next lLoop1
  Loop Until lHValue = LBound(vArray)
End Sub

Public Function OpenThisFile(stFile As String, lShowHow As Long, sParams As String, LhWnd As Long) As Variant
' Simply attempts to open a file with the Shell command, if errors are encountered then...
'   > tries to open it with an API call using extensions to find associated executables, if error then...
'   > prompts user with the "Open With ..." routine
Dim lRet As Long, stRet As String, ErrID As Long

On Error GoTo TryAPIcall
    lRet = -1   ' set default value -- meaning failure
    If Len(sParams) > 0 Then sParams = " " & sParams    ' if no optional parameters, then format with a space
    lRet = Shell(stFile & sParams, lShowHow)       ' attempt simple shell command
OpenThisFile = lRet
Exit Function

TryAPIcall:
Err.Clear
' if above shell function failed, then try an association open based on the file extension
    ErrID = apiShellExecute(LhWnd, "OPEN", _
            stFile, sParams, App.Path, lShowHow)
    ' Errors will be a retruned value of <32
    If ErrID < 32& Then
        Select Case ErrID
            Case 31&:
                'Try the OpenWith dialog
                lRet = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " _
                        & stFile, 1)
            Case 0&:
                stRet = "Error: Out of Memory/Resources. Couldn't Execute!"
            Case 2&:
                stRet = "Error: File not found.  Couldn't Execute!"
            Case 3&:
                stRet = "Error: Path not found. Couldn't Execute!"
            Case 11&:
                stRet = "Error:  Bad File Format. Couldn't Execute!"
            Case Else:
        End Select
        If ErrID <> 31 Then
            lRet = -1 ' failure
            MsgBox stRet, vbExclamation + vbOKOnly  ' display error
        End If
        OpenThisFile = lRet
    Else
        lRet = 69
    End If
Resume Next
End Function

Public Function StripFile(Pathname As String, DPNEm As String) As String
Dim ChrsIn As String, ChrsOut As String, IdX As Integer, Chrs As Integer

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo StripFile_General_ErrTrap
If Pathname = "" Then Exit Function
ChrsIn = Pathname
Select Case InStr("DPNEm", DPNEm)
Case 1:     ' Return the Drive Letter
    GoSub ExtractDrive
Case 2:     ' Return the Path
    GoSub ExtractPath
Case 3:     ' Return the File Name
    GoSub ExtractName
Case 4:     ' Return the File Extension
    GoSub ExtractExtension
Case 5:     ' Return filename less the extension
    GoSub ExtractName
    ChrsIn = StripFile
    GoSub ExtractExtension
    StripFile = Left(ChrsIn, Chrs - 1)
End Select
Exit Function

ExtractDrive:
Chrs = InStr(ChrsIn, ":\") 'check to see if a forward slash exists
If Chrs Then 'if a forward slash is found in the passed string
   ChrsOut = Left(ChrsIn, Chrs + 1) 'get the drive
End If
StripFile = ChrsOut 'return the drive to the user
Return
 
ExtractExtension:
Chrs = InStr(ChrsIn, ".") 'check to see if a full stop exists
If Chrs Then 'if a full stop is found in the passed string
IdX = Chrs
Do While IdX > 0
    IdX = InStr(IdX + 1, ChrsIn, ".")
    If IdX Then Chrs = IdX
Loop
   ChrsOut = Mid(ChrsIn, Chrs + 1) 'get the extension
Else
    ChrsOut = ""
End If
StripFile = ChrsOut 'return the extension to the user
Return

ExtractName:
If InStr(ChrsIn, "\") Then 'check to see if a forward slash exists
   For IdX = Len(ChrsIn) To 1 Step -1 'step though until full name is extracted
       If Mid(ChrsIn, IdX, 1) = "\" Then
          ChrsOut = Mid(ChrsIn, IdX + 1)
          Exit For
       End If
   Next IdX
ElseIf InStr(ChrsIn, ":") = 2 Then 'otherwise, check to see if a colon exists
   ChrsOut = Mid(ChrsIn, 3)        'if so, return the filename
Else
   ChrsOut = ChrsIn 'otherwise, return the original string
End If
StripFile = ChrsOut 'return the filename to the user
Return

ExtractPath:
If InStr(ChrsIn, "\") Then 'check to see if a forward slash exists
   For IdX = Len(ChrsIn) To 1 Step -1 'step though until full name is extracted
       If Mid(ChrsIn, IdX, 1) = "\" Then
          ChrsOut = Left(ChrsIn, IdX)
          Exit For
       End If
   Next IdX
ElseIf InStr(ChrsIn, ":") = 2 Then 'otherwise, check to see if a colon exists
   ChrsOut = CurDir(ChrsIn)
   If Len(ChrsOut) = 0 Then
      ChrsOut = CurDir
   End If
Else
   ChrsOut = CurDir 'otherwise, return the current directory
End If
StripFile = ChrsOut 'return the filenames path to the user
Return
Exit Function

' Inserted by LaVolpe OnError Insertion Program.
StripFile_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: StripFile" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Function

Public Sub ExtractFontInfo(vFontInfo As Variant, bColor As Boolean, Optional bSave As Boolean = False)
Dim sFontInfo As String, I As Integer
If bSave Then
    If bColor Then
        SaveSetting "LVripper", "Defaults", "FontColor", CStr(vFontInfo)
    Else
        For I = 0 To UBound(vFontInfo)
            sFontInfo = sFontInfo & vFontInfo(I) & ";"
        Next
        SaveSetting "LVripper", "Defaults", "FontInfo", sFontInfo
    End If
    Dim lLastPos(0 To 1) As Long
    For I = 0 To Forms.Count - 1
        If Forms(I).hWnd = frmParent.hWnd Then
            With Forms(I).rtfView
                If bColor Then
                    lLastPos(0) = .SelStart
                    lLastPos(1) = .SelLength
                    .HideSelection = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SelColor = vFontInfo
                    .HideSelection = False
                    .SelStart = lLastPos(0)
                    .SelLength = lLastPos(1)
                Else
                    .Font.Name = vFontInfo(0)
                    .Font.Size = vFontInfo(1)
                    .Font.Bold = vFontInfo(2)
                    .Font.Italic = vFontInfo(3)
                    .Font.Underline = vFontInfo(4)
                    .Font.Strikethrough = vFontInfo(5)
                End If
            End With
        End If
    Next
Else
    If bColor Then
        vFontInfo = GetSetting("LVripper", "Defaults", "FontColor", "0")
    Else
        Dim J As Integer
        sFontInfo = GetSetting("LVripper", "Defaults", "FontInfo", "Times New Roman;11.25;False;False;False;False;")
        ReDim vFontInfo(0 To 5)
        For I = 0 To UBound(vFontInfo)
           J = InStr(sFontInfo, ";")
           vFontInfo(I) = Left(sFontInfo, J - 1)
           sFontInfo = Mid(sFontInfo, J + 1)
        Next
    End If
End If
End Sub

Public Function FindListItem(ObjectHwnd As Long, bListBox As Boolean, bExactMatch As Boolean, sCriteria As String) As Long
' Function checks listbox contents for match of sCriteria if bListBox = True, otherwise
' checks combobox contents for match of sCriteria if bListBox = False
Dim lMatchType As Long
If bListBox = True Then
    If bExactMatch = False Then lMatchType = LB_FINDSTRING Else lMatchType = LB_FINDSTRINGEXACT
Else
    If bExactMatch = False Then lMatchType = CB_FINDSTRING Else lMatchType = CB_FINDSTRINGEXACT
End If
FindListItem = SendMessageLong(ObjectHwnd, lMatchType, -1, ByVal sCriteria)
End Function

