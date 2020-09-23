VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C8D0D4&
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   7215
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   5640
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2752
            Text            =   "Length to Beginning "
            TextSave        =   "Length to Beginning "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Amount Read"
            TextSave        =   "Amount Read"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Length to End"
            TextSave        =   "Length to End"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Read Now Using Above Criteria"
      Height          =   315
      Left            =   2820
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   405
      Width           =   4365
   End
   Begin VB.ComboBox cboAnchor 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   5520
      List            =   "Form1.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   60
      Width           =   1665
   End
   Begin VB.ComboBox cboDirection 
      Height          =   315
      ItemData        =   "Form1.frx":0043
      Left            =   2820
      List            =   "Form1.frx":004D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   1260
   End
   Begin VB.ComboBox cboPct 
      Height          =   315
      ItemData        =   "Form1.frx":0061
      Left            =   4110
      List            =   "Form1.frx":007A
      TabIndex        =   2
      Text            =   "5%"
      Top             =   60
      Width           =   975
   End
   Begin VB.ComboBox cboMode 
      Height          =   315
      ItemData        =   "Form1.frx":00A1
      Left            =   690
      List            =   "Form1.frx":00AB
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   1590
   End
   Begin RichTextLib.RichTextBox rtfView 
      Height          =   4920
      Left            =   30
      TabIndex        =   6
      Top             =   690
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   8678
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"Form1.frx":00C8
   End
   Begin VB.ComboBox cboChunk 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":014A
      Left            =   4110
      List            =   "Form1.frx":0166
      TabIndex        =   3
      Text            =   "64 mb"
      Top             =   60
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C8D0D4&
      Caption         =   "Right click below for more options...."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   450
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C8D0D4&
      Caption         =   "from"
      Height          =   195
      Index           =   2
      Left            =   5175
      TabIndex        =   9
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C8D0D4&
      Caption         =   "Read"
      Height          =   195
      Index           =   1
      Left            =   2340
      TabIndex        =   8
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C8D0D4&
      Caption         =   "MODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Private sSource As String, sDest As String
Private Type FilePosDat
    Start As Long
    Stop As Long
End Type
Private FilePos As FilePosDat
Private lSearchDat(1 To 2) As Long

Private Sub cboAnchor_Click()
' depending on whether moving thru file using Next/Previous option, set the direction appropriately
If cboAnchor.ListIndex = 2 Then cboDirection.ListIndex = 1
If cboAnchor.ListIndex < 1 Then cboDirection.ListIndex = 0
End Sub

Private Sub cboChunk_Validate(Cancel As Boolean)
' This routines simply helps ensure the file chunk size entered is a valid entry
If cboChunk.ListIndex > -1 Then Exit Sub
Dim sCurVals As String
If Val(cboChunk.Text) < 1 Then     ' string or negative provided
    sCurVals = cboChunk.Text
    If Trim(sCurVals) = "" Then sCurVals = "{blank}"
    MsgBox "Provide a file chunk size to read from the source file." & vbCrLf & "The entry " & sCurVals & " isn't valid.", vbInformation + vbOKOnly
    cboChunk.ListIndex = 1
    Exit Sub
End If
' MB or KB?
If InStr(cboChunk.Text, "m") = 0 And InStr(cboChunk.Text, "k") = 0 Then
    If Val(cboChunk.Text) > 5 Then cboChunk.Text = cboChunk.Text & "kb" Else cboChunk.Text = cboChunk.Text & "mb"
End If
Dim I As Integer, vList() As Variant, X As Integer, Y As Long, J As Integer, sTmp As String
' let's remove any non-numerical characters
sTmp = cboChunk.Text
For I = Len(cboChunk.Text) To 1 Step -1
    If InStr("0123456789kbm", Mid$(sTmp, I, 1)) = 0 Then
        sTmp = Replace$(sTmp, Mid$(sTmp, I, 1), "")
    End If
Next
cboChunk.Text = sTmp
' get location of the KB or MB
X = InStr(cboChunk.Text, "m")
Y = 2
If X = 0 Then
    ' if MB wasn't provided, assume KB
    X = InStr(cboChunk, "k")
    Y = 1
End If
' now format it the way we want # KB or # MB
cboChunk.Text = Left(cboChunk.Text, X - 1) & " " & Choose(Y, "kb", "mb")
' let's see if the formatted entry is already in the list
ReDim vList(0 To cboChunk.ListCount - 1)
For I = 0 To cboChunk.ListCount - 1
    sCurVals = sCurVals & "|" & cboChunk.List(I)
    If cboChunk.ItemData(I) = 1 Then vList(I) = "kb" Else vList(I) = "mb"
    ' we'll build a list to sort by
    vList(I) = vList(I) & Format(Val(cboChunk.List(I)), "0000000000.0000000000")
    ' keep track of the listindex for the text entered
    If cboChunk.List(I) = cboChunk.Text Then X = I
Next
If InStr(sCurVals, cboChunk.Text) = 0 Then
    ' new entry does not exist in current list
    ReDim Preserve vList(0 To cboChunk.ListCount)
    ' let's add the new entry to the array so we can sort it
    sCurVals = cboChunk.Text
    ' we put the mb or kb in the beginning of the string array. This way 1mb falls after 256kb
    vList(cboChunk.ListCount) = Choose(Y, "kb", "mb") & Format(Val(sCurVals), "0000000000.0000000000")
    ShellSortArray vList
    ' we're going to repopulate the listing in numerical order
    cboChunk.Clear
    For I = 0 To UBound(vList)
        ' add sorted items to the combo box
        sTmp = CStr(Val(Mid(vList(I), 3))) & " "
        cboChunk.AddItem sTmp & Left(vList(I), 2)
        If InStr(vList(I), "kb") Then cboChunk.ItemData(I) = 1 Else cboChunk.ItemData(I) = 1024
        ' keep track of the listindex for the text entered
        If cboChunk.List(I) = sCurVals Then X = I
    Next
End If
cboChunk.ListIndex = X
Erase vList
End Sub

Private Sub cboMode_Click()
' enable/disable comboboxes dependent on the mode selected
cboPct.Enabled = CBool(cboMode.ListIndex - 1)
cboPct.Visible = CBool(cboMode.ListIndex - 1)
cboChunk.Visible = CBool(cboMode.ListIndex)
cboChunk.Enabled = CBool(cboMode.ListIndex)
End Sub

Private Sub cboPct_Validate(Cancel As Boolean)
' let's make sure the entry provided is valid
If cboPct.ListIndex > -1 Then Exit Sub
Dim sCurVals As String
If Val(cboPct.Text) < 1 Then
    sCurVals = cboPct.Text
    If Trim(sCurVals) = "" Then sCurVals = "{blank}"
    MsgBox "Provide a percentage to read from the source file." & vbCrLf & "The entry " & sCurVals & " isn't valid.", vbInformation + vbOKOnly
    cboPct.ListIndex = 1
    Exit Sub
End If
Dim I As Integer, vList() As Variant, X As Integer, sTmp As String
' let's remove any non-numerical characters
sTmp = cboPct.Text
For I = Len(cboPct.Text) To 1 Step -1
    If InStr("0123456789", Mid$(sTmp, I, 1)) = 0 Then
        sTmp = Replace$(sTmp, Mid$(sTmp, I, 1), "")
    End If
Next
cboPct.Text = sTmp & "%"
ReDim vList(0 To cboPct.ListCount - 1)
' let's see if the formatted entry is already in the list
For I = 0 To cboPct.ListCount - 1
    ' we'll build a list to sort by
    sCurVals = sCurVals & "|" & cboPct.List(I)
    vList(I) = Val(cboPct.List(I))
    ' keep track of the listindex for the text entered
    If cboPct.List(I) = cboPct.Text Then X = I
Next
If InStr(sCurVals, cboPct.Text) = 0 Then
    ' new entry does not exist in current list
    ReDim Preserve vList(0 To cboPct.ListCount)
    ' let's add the new entry to the array so we can sort it
    sCurVals = cboPct.Text
    ' we put the mb or kb in the beginning of the string array. This way 1mb falls after 256kb
    vList(cboPct.ListCount) = Val(sCurVals)
    ShellSortArray vList
    ' we're going to repopulate the listing in numerical order
    cboPct.Clear
    For I = 0 To UBound(vList)
        ' add sorted items to the combo box
        cboPct.AddItem CStr(vList(I)) & "%"
        cboPct.ItemData(I) = 1
        ' keep track of the listindex for the text entered
        If cboPct.List(I) = sCurVals Then X = I
    Next
End If
cboPct.ListIndex = X
Erase vList
End Sub

Private Sub cmdRead_Click()
' Read another section of the file, if possible
If cboAnchor.ListIndex = 1 And ((FilePos.Start = 0 And cboDirection.ListIndex = 1) Or (FilePos.Stop < 0 And cboDirection.ListIndex = 0)) Then
    MsgBox "Can't read the file with that criteria. The file is already at the " & _
        Choose(cboDirection.ListIndex + 1, "end", "beginning") & ".", vbInformation + vbOKOnly
    Exit Sub
End If
Dim lEndPos As Long
' calculate the next pct/chunk to read
If cboPct.Enabled = True Then
    lEndPos = Val(cboPct) * cboPct.ItemData(cboPct.ListIndex)
Else
    lEndPos = Val(cboChunk) * (cboChunk.ItemData(cboChunk.ListIndex) * 1024#)
End If
' set the new start position to where the stop position was
If cboDirection.ListIndex = 0 Then FilePos.Start = FilePos.Stop
If cboAnchor.ListIndex = 0 Then FilePos.Start = 0
MousePointer = vbHourglass
rtfView.FileName = ""
' call function to read the file which places it in a temp file
ReadFilePart sSource, sDest, FilePos.Start, FilePos.Stop, lEndPos, cboMode.ListIndex, cboDirection.ListIndex, cboAnchor.ListIndex

' now show that file
rtfView.FileName = sDest
If Val(GetSetting("LVripper", "Defaults", "FontColor", "0")) Then
    With rtfView
        .HideSelection = True
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelColor = GetSetting("LVripper", "Defaults", "FontColor", "0")
        .SelStart = 0
        .HideSelection = False
    End With
End If
MousePointer = vbDefault
' move the anchor option from Beginning/Last to Current Position
If cboAnchor.ListIndex <> 1 Then cboAnchor.ListIndex = 1
UpdateStatusBar
End Sub

Private Sub Form_Activate()
' update the WordWrap checkmark on MDI form
With frmParent
    .mnuChild(8).Checked = CBool(Tag)
    .mnuOpts(5).Checked = .mnuChild(8).Checked
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbCtrlMask Then
    Select Case KeyCode
    Case vbKeyF     ' Find
        KeyCode = 0
        frmParent.picSearch.Visible = True
        frmParent.txtFind.SetFocus
        SetSearchOptions True
    Case vbKeyO ' Open
        KeyCode = 0
        Call frmParent.mnuFile_Click(0)
    Case vbKeyF4    ' Close
        KeyCode = 0
        Unload Me
    End Select
Else
    If Shift = vbAltMask Then
        Select Case KeyCode
        Case vbKeyF    ' Find first
            KeyCode = 0
            Call frmParent.mnuSearch_Click(2)
        Case vbKeyN ' Find Next
            KeyCode = 0
            Call frmParent.mnuSearch_Click(3)
        Case vbKeyP ' Find previous
            KeyCode = 0
            Call frmParent.mnuSearch_Click(4)
        Case vbKeyL ' Find Last
            KeyCode = 0
            Call frmParent.mnuSearch_Click(5)
        End Select
    Else
        If Shift = vbCtrlMask + vbShiftMask And KeyCode = vbKeyF4 Then
            KeyCode = 0 ' Close All
            Call frmParent.mnuFile_Click(2)
        End If
    End If
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
cboAnchor.ListIndex = 0         ' Beginning of file
cboDirection.ListIndex = 0      ' Next vs Previous
cboChunk.ListIndex = 2          ' 64kb
cboPct.ListIndex = 0              ' 5%
' set the mode passed off of stored preference
If frmParent.mnuMode(0).Checked Then cboMode.ListIndex = 0 Else cboMode.ListIndex = 1
sDest = GetTempFile(True)    ' get a temp file name to use & ID source file name
sSource = frmParent.dlgCommon.FileName
Caption = sSource                ' update caption bar with file name
Tag = "1"                              ' wordwrap option
Dim vFontInfo() As Variant      ' set RTF box with preferred font
ExtractFontInfo vFontInfo, False
With rtfView
    .Font.Name = vFontInfo(0)
    .Font.Size = vFontInfo(1)
    .Font.Bold = vFontInfo(2)
    .Font.Italic = vFontInfo(3)
    .Font.Underline = vFontInfo(4)
    .Font.Strikethrough = vFontInfo(5)
End With
Call cmdRead_Click
WindowState = vbMaximized
frmParent.mnuSearch(0).Enabled = True
End Sub

Private Sub UpdateStatusBar()
' updates the status bar with numbers
Dim lSize(0 To 1) As Long, iKB As Long, lStart As Long, lHDL As Long
' get portion file size
lHDL = RetrieveFileHandle(sDest)
lSize(0) = GetFileSize(lHDL, 0)
CloseHandle lHDL
' get original file size
lHDL = RetrieveFileHandle(sSource)
lSize(1) = GetFileSize(lHDL, 0)
CloseHandle lHDL
' byte ref to end of portion read
If FilePos.Stop < 0 Then lStart = lSize(1) Else lStart = FilePos.Stop
' now calculate where the file section started reading at & format appropriately
lStart = lStart - lSize(0)
If lStart < 1024000 Then iKB = 1024 Else iKB = 1024000
sBar.Panels(2) = Format(lStart / iKB, "0.#") & " " & IIf(iKB = 1024, "kb", "mb")
' calculate bytes to end of file
If FilePos.Stop < 0 Then lStart = lSize(1) Else lStart = FilePos.Stop
lStart = lSize(1) - lStart
If lStart < 1024000 Then iKB = 1024 Else iKB = 1024000
sBar.Panels(6) = Format(lStart / iKB, "0.#") & " " & IIf(iKB = 1024, "kb", "mb")
' calculate number of bytes read
If lSize(0) < 1024000 Then iKB = 1024 Else iKB = 1024000
sBar.Panels(4) = Format(lSize(0) / iKB, "0.#") & " " & IIf(iKB = 1024, "kb", "mb")
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then Exit Sub
' expand the text box to match actual window size
On Error Resume Next
rtfView.Height = Height - 1440
rtfView.Width = Width - 195
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
' delete the temporary file
On Error Resume Next
If Len(Dir(sDest)) Then DeleteFile sDest
SetSearchOptions False
End Sub

Private Sub rtfView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu frmParent.mnuChildren
End Sub

Private Sub SetSearchOptions(bOnOff As Boolean)
Dim I As Integer
If bOnOff = False Then
    If Forms.Count < 3 Then
        ' hide and disable search options if no child forms are going to exist
        frmParent.picSearch.Visible = False
        frmParent.mnuSearch(0).Enabled = False
        For I = 2 To frmParent.mnuSearch.UBound
            frmParent.mnuSearch(I).Enabled = False
        Next
    End If
Else
    ' enable search options when a child form is displayed
    For I = 2 To frmParent.mnuSearch.UBound
        frmParent.mnuSearch(I).Enabled = True
    Next
End If
End Sub
