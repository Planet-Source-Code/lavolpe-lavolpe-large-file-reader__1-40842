VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmParent 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Large File Reader"
   ClientHeight    =   5085
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7305
   Icon            =   "frmParent.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   9330
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select File To Open"
   End
   Begin VB.PictureBox picSearch 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7245
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7305
      Begin VB.CommandButton Command1 
         Caption         =   "Stop Searching"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3930
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Close Search"
         Height          =   315
         Left            =   5505
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   300
         Width           =   1560
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&First"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   3945
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   775
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Next"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   775
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Previous"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   5505
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   775
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Last"
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   6285
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   775
      End
      Begin VB.ComboBox txtFind 
         Height          =   315
         Left            =   765
         TabIndex        =   1
         Top             =   0
         Width           =   3165
      End
      Begin VB.CheckBox chkWholeWord 
         Caption         =   "Whole words only?"
         Height          =   315
         Left            =   765
         TabIndex        =   2
         Top             =   300
         Width           =   1695
      End
      Begin VB.CheckBox chkCase 
         Caption         =   "Match case?"
         Height          =   315
         Left            =   2505
         TabIndex        =   3
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Find..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   11
         Top             =   135
         Width           =   675
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4710
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "File Ripper.  Drop files on this task bar to open them or select menu File | Open "
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Close Current"
         Index           =   1
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Close &ALL"
         Index           =   2
         Shortcut        =   +^{F4}
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   4
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Options"
      Index           =   1
      Begin VB.Menu mnuOpts 
         Caption         =   "&Font"
         Index           =   0
      End
      Begin VB.Menu mnuOpts 
         Caption         =   "&Font Color"
         Index           =   1
      End
      Begin VB.Menu mnuOpts 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuOpts 
         Caption         =   "Memory Chunks"
         Index           =   3
      End
      Begin VB.Menu mnuOpts 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuOpts 
         Caption         =   "Word Wrap"
         Checked         =   -1  'True
         Index           =   5
      End
      Begin VB.Menu mnuOpts 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuOpts 
         Caption         =   "Preferred Mode"
         Index           =   7
         Begin VB.Menu mnuMode 
            Caption         =   "By Percentage"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuMode 
            Caption         =   "By File Chunks"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Search"
      Index           =   2
      Begin VB.Menu mnuSearch 
         Caption         =   "&Show Search Window"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Find &First         Alt+F"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Find &Next        Alt+N"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Find &Previos    Alt+P"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Find &Last         Alt+L"
         Enabled         =   0   'False
         Index           =   5
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWinSub 
         Caption         =   "&Cascade"
         Index           =   0
      End
      Begin VB.Menu mnuWinSub 
         Caption         =   "Tile &Horizontally"
         Index           =   1
      End
      Begin VB.Menu mnuWinSub 
         Caption         =   "Tile &Vertically"
         Index           =   2
      End
   End
   Begin VB.Menu mnuChildren 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuChild 
         Caption         =   "&Copy Selected Text"
         Index           =   0
      End
      Begin VB.Menu mnuChild 
         Caption         =   "Copy &ENTIRE Text"
         Index           =   1
      End
      Begin VB.Menu mnuChild 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuChild 
         Caption         =   "&Paste"
         Index           =   3
      End
      Begin VB.Menu mnuChild 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuChild 
         Caption         =   "&Send to..."
         Index           =   5
         Begin VB.Menu mnuSendTo 
            Caption         =   "Other Application..."
            Index           =   0
         End
         Begin VB.Menu mnuSendTo 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuSendTo 
            Caption         =   "NotePad"
            Index           =   2
         End
         Begin VB.Menu mnuSendTo 
            Caption         =   "WordPad"
            Index           =   3
         End
      End
      Begin VB.Menu mnuChild 
         Caption         =   "Save &As"
         Index           =   6
         Begin VB.Menu mnuSaveAs 
            Caption         =   "ASCII Text"
            Index           =   0
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Rich Text Format"
            Index           =   1
         End
      End
      Begin VB.Menu mnuChild 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuChild 
         Caption         =   "WordWrap"
         Checked         =   -1  'True
         Index           =   8
      End
   End
End
Attribute VB_Name = "frmParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private StopSearch As Boolean

Private Sub cmdFind_Click(Index As Integer)
' A bit cumbersome, but it worked & really don't have the time to recode this search engine
' If you can't follow along, sorry.  The gist is to process a 1/4 million characters at a time
'   to speed up the search. Allows user to search for last occurrence along with first occurrence
With ActiveForm.rtfView
    .HideSelection = True
    StopSearch = False
    Dim lFoundLast(1 To 2)
    Dim iOptions As Integer, lStart As Long, iResult As Long, bFound As Boolean
    Dim prevWinPos As Integer, lLastPos As Long, sCriteria As String
    If chkWholeWord = 1 Then iOptions = 2
    If chkCase = 1 Then iOptions = iOptions + 4
    iOptions = iOptions + 8             ' option not to highlight matches
    On Error GoTo EndSearch
BeginSearch:
    lFoundLast(1) = .SelStart
    lFoundLast(2) = .SelLength
    .SelLength = 0
    Command1.Enabled = True
        ' set starting point in text document for comparison
        lStart = Choose(Index + 1, 0, .SelStart + 1, .SelStart - 1, .SelStart + 1)
        Select Case Index
        Case 0, 1   ' Find first, find next
            ' conduct search
            iResult = .Find(txtFind, lStart, , iOptions)
        Case 2  ' Find previous
            ' Since the Find function won't go backwards, need to do a little manipulation
            Command1.Enabled = True ' enable an abort
            lLastPos = lStart - 250000            ' go back in the text document, 5000 characters
FindPreviousMatch:
            If lLastPos < 0 Then lLastPos = 0   ' if gone back too far go back to zero
            DoEvents
            iResult = .Find(txtFind, lLastPos, , iOptions)  ' Conduct a search from 'back' position
            If iResult > -1 And iResult < lStart Then  ' If match found which is not current text position, then
                Do Until iResult = -1                   ' Search again to ensure a later match isn't found
                    DoEvents                            ' Keep track of the last match position
                    lFoundLast(1) = iResult: lFoundLast(2) = Len(txtFind.Text)
                    iResult = .Find(txtFind, iResult + 1, lStart - 1, iOptions) ' Conduct next search
                    If StopSearch = True Then Exit Do   ' User aborts search if variable = True
                Loop
                iResult = lFoundLast(1)                 ' Reset the iResult variable to last match position
            Else    ' No match found 50000 characters back, so go another 5000 back and try again
                If lLastPos > 0 Then    ' however, if already went back to begining, abort here
                    lLastPos = lLastPos - 250000  ' reset another 50,000 characters
                    If StopSearch = False Then GoTo FindPreviousMatch ' try search again
                Else    ' went back to begining with no match, reset variable and exit out
                    iResult = -1
                End If
            End If
        Case 3  ' Find last possible match
            .SelStart = Len(.Text)
            Index = 2: GoTo BeginSearch     ' Move to end of file and conduct a Previous Search
            Command1.Enabled = True ' enable an abort
            iResult = .Find(txtFind, lStart, , iOptions)    ' Conduct search
            Do While iResult <> -1
                DoEvents                ' Continue search until end of document
                bFound = True           ' keep track of any matches
                lFoundLast(1) = iResult: lFoundLast(2) = Len(txtFind.Text)
                                        ' conduct next search
                iResult = .Find(txtFind, iResult + 1, , iOptions)
                If StopSearch = True Then Exit Do   ' Abort if user chose to
            Loop
        End Select
    Command1.Enabled = False
    If StopSearch = False Then
        ' Display results
        ' If finding first, next or previous & no match found, or
        '   trying to find last and no match found then display error message
        If (Index < 3 And iResult = -1) Or (Index = 3 And bFound = False) Then
            .SelStart = lFoundLast(1): .SelLength = lFoundLast(2)
            MsgBox Chr(34) & txtFind & Chr(34) & " not found", vbInformation + vbOKOnly
        End If
        ' Now, if match found save position and show results
        If Index < 2 Then   ' attempted Find first, next or previous
            If iResult > -1 Then    ' match found
                cmdFind(1).Default = True   ' set default to Find Next & save position of match
                cmdFind(1).SetFocus
                lFoundLast(1) = iResult: lFoundLast(2) = Len(txtFind.Text)
            End If
        Else    ' Finding last match
            cmdFind(2).Default = True   ' Set find previous as default button
            cmdFind(2).SetFocus
            If bFound = True Then       ' Match found, save position of match
                .SelStart = lFoundLast(1)
                .SelLength = lFoundLast(2)
            End If
        End If
    End If
    StopSearch = False               ' Reset variable since this form doesn't close
    ' highlight match position or starting position, if no match found, in the text document
    .HideSelection = False
    .SelStart = lFoundLast(1): .SelLength = lFoundLast(2)
    Me.SetFocus
End With
Exit Sub

EndSearch:  ' Error during search
ActiveForm.rtfView.HideSelection = False
ActiveForm.rtfView.SelStart = lFoundLast(1): ActiveForm.rtfView.SelLength = lFoundLast(2)
End Sub

Private Sub MDIForm_Load()
' get amount of memory to use for ripping thru files
lReadSize = Val(GetSetting("LVripper", "Defaults", "MemoryChunk", "32000"))
' get preferred mode (Pct or Chunk) & update checkmarks
mnuMode(0).Checked = CBool(Val(GetSetting("LVripper", "Defaults", "Pref", "0")) = 0)
mnuMode(1).Checked = Not mnuMode(0).Checked
If Len(Command()) Then  ' opened from a command line
    Dim sFile As String
    sFile = Replace(Command(), Chr(34), "")
    If Len(Dir(sFile)) Then
        Show
        dlgCommon.FileName = sFile
        Call mnuFile_Click(-1)
    End If
End If
End Sub

Private Sub MDIForm_Terminate()
Unload Me
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
CloseAllViewers
Dim sTempoaryFiles As String
sTempoaryFiles = GetTempFile(True)
sTempoaryFiles = StripFile(sTempoaryFiles, "P") & "~Lv*.tmp"
On Error Resume Next
Kill sTempoaryFiles
End Sub

Private Sub mnuChild_Click(Index As Integer)
If Index > 3 And Index < 8 Then Exit Sub
Dim frmChild As Form
On Error GoTo FailedClickOption
Set frmChild = ActiveForm
With frmChild
    Select Case Index
    Case 0: ' Copy selected text
        If Len(.rtfView.SelText) = 0 Then
            MsgBox "No text has been selected", vbInformation + vbOKOnly
        Else
            Clipboard.Clear
            Clipboard.SetText .rtfView.SelText
            MsgBox "Text Copied", vbInformation + vbOKOnly
        End If
    Case 1: ' Copy all text
        Clipboard.Clear
        Clipboard.SetText .rtfView.Text
        MsgBox "Text Copied", vbInformation + vbOKOnly
    Case 3: ' Paste
        On Error GoTo NoPaste
        .rtfView.SelText = Clipboard.GetText
    Case 8: ' WordWrap
        Call mnuOpts_Click(5)
    End Select
End With
Set frmChild = Nothing
Exit Sub

FailedClickOption:
MsgBox "<Error> " & Err.Description, vbInformation + vbOKOnly
Set frmChild = Nothing
Exit Sub
NoPaste:
MsgBox "Couldn't paste what was in the clipboard. Re-copy the text & try to paste again.", vbInformation + vbOKOnly
Resume Next
End Sub

Friend Sub mnuFile_Click(Index As Integer)
Select Case Index
Case 0, -1:    ' New File
    On Error GoTo UserCnx
    If Index = 0 Then
        With dlgCommon
            .Flags = cdlOFNFileMustExist
            .FileName = ""
            .DialogTitle = "Select File to Open"
        End With
        dlgCommon.ShowOpen
    End If
    Dim fNewView As New Form1
    fNewView.Show
Case 1
    On Error Resume Next
    Unload ActiveForm
Case 2
    CloseAllViewers
Case 4
    CloseAllViewers
    End
End Select

UserCnx:
End Sub

Private Sub CloseAllViewers()
On Error Resume Next
Dim I As Integer
For I = Forms.Count - 1 To 0 Step -1
    If Forms(I).hWnd <> hWnd Then Unload Forms(I)
Next
End Sub

Private Sub mnuMode_Click(Index As Integer)
' update mode preference
If mnuMode(Index).Checked Then Exit Sub
mnuMode(Index).Checked = True
mnuMode(Abs(Index - 1)).Checked = False
SaveSetting "LVripper", "Defaults", "Pref", CStr(Index)
End Sub

Private Sub mnuOpts_Click(Index As Integer)
Dim vFontColor As Variant
Select Case Index
Case 0: ' Font
    Dim vFontInfo() As Variant
    ExtractFontInfo vFontInfo, False
    On Error GoTo UserCnx
    With dlgCommon
        .Flags = cdlCFBoth
        .FontName = vFontInfo(0)
        .FontSize = vFontInfo(1)
        .FontBold = vFontInfo(2)
        .FontItalic = vFontInfo(3)
        .FontUnderline = vFontInfo(4)
        .FontStrikethru = vFontInfo(5)
    End With
    dlgCommon.ShowFont
    With dlgCommon
        vFontInfo(0) = .FontName
        vFontInfo(1) = .FontSize
        vFontInfo(2) = .FontBold
        vFontInfo(3) = .FontItalic
        vFontInfo(4) = .FontUnderline
        vFontInfo(5) = .FontStrikethru
    End With
    ExtractFontInfo vFontInfo, False, True
    Erase vFontInfo
Case 1: ' Color
    ExtractFontInfo vFontColor, True
    On Error GoTo UserCnx
    dlgCommon.Flags = cdlCCPreventFullOpen Or cdlCCRGBInit
    dlgCommon.Color = vFontColor
    dlgCommon.ShowColor
    vFontColor = dlgCommon.Color
    ExtractFontInfo vFontColor, True, True
Case 3: ' Memory
    Dim sMsg As String, sValue As String
    sMsg = "How much memory do you want to use. The larger amount, the faster the program can rip through your files." _
        & vbCrLf & "NOTE: Allocating too much memory may cause Out of Memory errors." & vbCrLf & vbCrLf _
        & "Enter the value in kilobytes (i.e., 32 for 32 kb)"
    sValue = InputBox(sMsg, "Memory Chunks", CStr(lReadSize / 1000))
    If Val(sValue) > 0 Then
        lReadSize = Val(sValue) * 1000
        SaveSetting "LVripper", "Defaults", "MemoryChunk", CStr(lReadSize)
    End If
Case 5  ' WordWrap
    If Forms.Count = 1 Then Exit Sub
    ActiveForm.Tag = CStr(CInt(mnuOpts(5).Checked) + 1)
    mnuOpts(5).Checked = CBool(ActiveForm.Tag)
    mnuChild(8).Checked = mnuOpts(5).Checked
    SendMessageLong ActiveForm.rtfView.hWnd, EM_SETTARGETDEVICE, 0, ByVal CLng(Abs(CInt(mnuOpts(5).Checked) + 1))
End Select
UserCnx:
End Sub

Private Sub mnuSaveAs_Click(Index As Integer)
On Error GoTo UserCnx
With dlgCommon
    .DefaultExt = Choose(Index + 1, rtfText, rtfRTF)
    .Filter = "Text Files|*.txt|Rich Text Files|*.rtf"
    .FilterIndex = Index + 1
    .DialogTitle = "Save As..."
    .InitDir = StripFile(ActiveForm.Caption, "P")
    .Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
    .FileName = ""
End With
dlgCommon.ShowSave
On Error GoTo FailedSave
ActiveForm.rtfView.SaveFile dlgCommon.FileName, Choose(Index + 1, rtfText, rtfRTF)
dlgCommon.DefaultExt = ""
Exit Sub

FailedSave:
MsgBox "File was not saved..." & vbCrLf & Err.Description, vbOKOnly
UserCnx:
dlgCommon.DefaultExt = ""
End Sub

Friend Sub mnuSearch_Click(Index As Integer)
Select Case Index
    Case 0
        If picSearch.Visible = False Then
            picSearch.Visible = True
            Dim I As Integer
            For I = 2 To frmParent.mnuSearch.UBound
                frmParent.mnuSearch(I).Enabled = True
            Next
        End If
        txtFind.SetFocus
    Case 2
        Call cmdFind_Click(0)
    Case 3
        Call cmdFind_Click(1)
    Case 4
        Call cmdFind_Click(2)
    Case 5
        Call cmdFind_Click(3)
End Select
End Sub

Private Sub mnuSendTo_Click(Index As Integer)
Dim sApplication As String, sTempFile As String, iOutType As Integer
iOutType = rtfRTF
Select Case Index
Case 0: ' Choose application
    On Error GoTo UserCnx
    With dlgCommon
        .DialogTitle = "Open with Which Application?"
        .Flags = cdlOFNFileMustExist
        .FileName = ""
        .Filter = "Executables|*.exe"
    End With
    dlgCommon.ShowOpen
    Load mnuSendTo(mnuSendTo.UBound + 1)
    mnuSendTo(mnuSendTo.UBound).Caption = StripFile(dlgCommon.FileName, "m")
    mnuSendTo(mnuSendTo.UBound).Tag = dlgCommon.FileName
    sApplication = dlgCommon.FileName
Case 2: ' NotePad
    sApplication = "Notepad.exe"
    iOutType = rtfText
Case 3: ' WordPad
    sApplication = "Wordpad.exe"
Case Else: ' Other
    sApplication = mnuSendTo(Index).Tag
End Select
sTempFile = GetTempFile(False)
ActiveForm.rtfView.SaveFile sTempFile, iOutType
OpenThisFile sApplication, 1, sTempFile, hWnd
UserCnx:
End Sub

Private Sub mnuWinSub_Click(Index As Integer)
If Index > 3 Then Exit Sub
Arrange Choose(Index + 1, vbCascade, vbTileHorizontal, vbTileVertical)
End Sub

Private Sub StatusBar1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(15) Then
' dropping file(s) on status bar
    On Error GoTo NoDrop
    Dim I As Integer
    With Data.Files
    For I = 1 To .Count
        dlgCommon.FileName = .Item(I)
        Call mnuFile_Click(-1)
    Next
    If .Count > 1 Then Arrange vbCascade
    End With
End If
Exit Sub

NoDrop:
MsgBox "Couldn't load all files." & vbcflf & Err.Description, vbInformation + vbOKOnly
End Sub

Private Sub txtFind_Change()
' enable/disable buttons depending on whether or not user entered text in the criteria box
cmdFind(0).Enabled = (Len(txtFind.Text) > 0)
cmdFind(1).Enabled = (Len(txtFind.Text) > 0)
cmdFind(2).Enabled = (Len(txtFind.Text) > 0)
cmdFind(3).Enabled = (Len(txtFind.Text) > 0)
cmdFind(0).Default = True
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
KeyCode = KeyCode
End Sub

Private Sub txtFind_Validate(Cancel As Boolean)
' save previous search criteria entries in combo box
If Len(txtFind.Text) = 0 Then Exit Sub
If FindListItem(txtFind.hWnd, False, True, txtFind.Text) < 0 Then txtFind.AddItem txtFind.Text
End Sub

Private Sub Command1_Click()
' Abort search
StopSearch = True
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
' Close search window
ActiveForm.rtfView.SetFocus
ActiveForm.rtfView.HideSelection = True
picSearch.Visible = False
Dim I As Integer
For I = 2 To frmParent.mnuSearch.UBound
    frmParent.mnuSearch(I).Enabled = False
Next
End Sub

