VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWalkthroughs 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Walkthroughs"
   ClientHeight    =   6516
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8772
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWalkthroughs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6516
   ScaleWidth      =   8772
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Prg1 
      Height          =   252
      Left            =   2040
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   2412
      _ExtentX        =   4255
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin RichTextLib.RichTextBox TxtSearch 
      Height          =   2172
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   3831
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmWalkthroughs.frx":030A
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   264
      Left            =   0
      TabIndex        =   0
      Top             =   6252
      Width           =   8772
      _ExtentX        =   15473
      _ExtentY        =   466
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15127
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2520
      Top             =   4560
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.ListBox lstWalkthru 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2352
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   3732
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6248
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   8772
      _ExtentX        =   15473
      _ExtentY        =   11028
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Platform"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Genre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   7850
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Size(KB)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "mnuOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print..."
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Search..."
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Add..."
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Edit..."
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Delete"
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Refresh"
         Index           =   9
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Remove &All"
         Index           =   10
      End
   End
End
Attribute VB_Name = "frmWalkthroughs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_USER = &H400
Private Const LVM_GETHEADER = (&H1000 + 31)
Private Const SB_GETRECT = (WM_USER + 10)
Private Const GWL_STYLE = (-16)
Private Const HDS_BUTTONS = &H2
Private FileData() As String
Private sData As String
Private bSel As Boolean

Private Sub Form_Load()
    Dim lHwnd As Long, lS As Long
    lHwnd = SendMessageByLong(ListView1.hWnd, LVM_GETHEADER, 0, 0)
    If (lHwnd <> 0) Then
        lS = GetWindowLong(lHwnd, GWL_STYLE)
        lS = lS And Not HDS_BUTTONS
        SetWindowLong lHwnd, GWL_STYLE, lS
    End If
    With frmWalkthroughs
        .Width = (Screen.Width) / 1.3
        .Height = (Screen.Height) / 1.3
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
    End With
    StatusBar1.Panels(1).Text = "Number of walkthroughs: 0"
    LoadWalkthroughs
    For j = 1 To ListView1.ListItems.Count
        ListView1.ListItems(j).Selected = False
    Next
    bSel = False
End Sub

Private Sub ShowProgress(cProgressBar As ProgressBar, cStatusBar As StatusBar, bShow As Boolean)

    Dim rc As RECT

    If bShow Then
        SendMessageAny cStatusBar.hWnd, SB_GETRECT, 0, rc

        With rc
            .Top = .Top * Screen.TwipsPerPixelY
            .Left = .Left * Screen.TwipsPerPixelX
            .Bottom = .Bottom * Screen.TwipsPerPixelY - .Top
            .Right = .Right * Screen.TwipsPerPixelX - .Left
        End With

        With cProgressBar
            SetParent .hWnd, cStatusBar.hWnd
            .Move rc.Left, rc.Top, rc.Right, rc.Bottom
            .Visible = True
        End With
    Else
        SetParent cProgressBar.hWnd, Me.hWnd
        cProgressBar.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    ListView1.Width = Me.ScaleWidth
    ListView1.Height = Me.ScaleHeight - StatusBar1.Height
    ListView1.ColumnHeaders(1).Width = ListView1.Width / 6
    ListView1.ColumnHeaders(2).Width = ListView1.Width / 6
    ListView1.ColumnHeaders(3).Width = ListView1.Width / 2
    ListView1.ColumnHeaders(4).Width = ListView1.Width / 6
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWalkthroughs
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count = 0 Then
        mnuInsert_Click
    Else
        If bSel = False Then mnuInsert_Click: Exit Sub
        mnuOpen_Click
    End If
End Sub

Private Sub RefreshList()
    Dim I As Integer
    For I = 1 To ListView1.ListItems.Count
        If Dir(lstWalkthru.List(I - 1)) = "" Then
            SetColour I, vbRed, True
            ListView1.ListItems(I).SubItems(3) = "N/A"
        Else
            SetColour I, vbBlack, False
            ListView1.ListItems(I).SubItems(3) = GetFileLength(lstWalkthru.List(I - 1))
        End If
    Next
    ListView1.Refresh
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If bSel = False Then Exit Sub
    If KeyCode = vbKeyDelete And ListView1.ListItems.Count <> 0 Then DeleteWalkthrough ListView1.SelectedItem.Index
    If KeyCode = vbKeyReturn Then ListView1_DblClick
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lstItem As ListItem
    Set lstItem = ListView1.HitTest(x, y)
    If lstItem Is Nothing Then
        For j = 1 To ListView1.ListItems.Count
            ListView1.ListItems(j).Selected = False
        Next
        bSel = False
    Else
        bSel = True
    End If
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And 2 Then
        If bSel = False Then GoTo notSelected
        For k = 0 To 10
            mnuFile(k).Visible = True
            mnuFile(k).Enabled = True
        Next
        mnuFile(10).Visible = False
        PopupMenu mnuOptions, , , , mnuFile(0)
    Else
        If bSel = False Then GoTo notSel
    End If
    Exit Sub
notSelected:
    For j = 1 To ListView1.ListItems.Count
        ListView1.ListItems(j).Selected = False
    Next
    If ListView1.ListItems.Count = 0 Then
        For k = 0 To 10
            mnuFile(k).Visible = True
            mnuFile(k).Enabled = False
        Next
        mnuFile(5).Enabled = True
        PopupMenu mnuOptions, , , , mnuFile(5)
    Else
        mnuFile(10).Visible = True
        mnuFile(9).Visible = True
        mnuFile(10).Enabled = True
        mnuFile(9).Enabled = True
        For k = 0 To 8
            mnuFile(k).Visible = False
            mnuFile(k).Enabled = False
        Next
        PopupMenu mnuOptions
    End If
    Exit Sub
notSel:
    For j = 1 To ListView1.ListItems.Count
        ListView1.ListItems(j).Selected = False
    Next
    Exit Sub
End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo vbError
    Dim Platform, Genre, Title, FilePath
    For I = 1 To Data.Files.Count
        frmWalkthroughs.SetFocus
        If (Effect And vbDropEffectCopy) Then
            FilePath = Data.Files.Item(I)
            If FilePath <> "" And Mid(FilePath, Len(FilePath) - 3, 1) = "." Then
                Platform = InputBoxEx("What platform is this game for?", "Platform", , , , , , , , "Courier New", "10")
                If Platform <> "" Then
                    Genre = InputBoxEx("What is the genre of this game?", "Genre", , , , , , , , "Courier New", "10")
                    If Genre <> "" Then
                        Title = InputBoxEx("What is the title of the game?", "Title", , , , , , , , "Courier New", "10")
                        If Title <> "" Then
                            AddWalkthrough Platform, Genre, Title, FilePath
                        End If
                    End If
                End If
            End If
        End If
    Next
vbError:
    Exit Sub
End Sub

Private Sub mnuDelete_Click()
    If ListView1.ListItems.Count <> 0 Then
        DeleteWalkthrough ListView1.SelectedItem.Index
    End If
End Sub

Private Function SearchFile(sFileName, sWord) As String
    Dim iFound, x, MatchCase
    If Right(sWord, 4) = " -mc" Then
        MatchCase = 4
        sWord = Left(sWord, Len(sWord) - 4)
    Else
        MatchCase = 0
    End If
    iFound = 0
    TxtSearch.LoadFile sFileName
    Me.MousePointer = 11
    Prg1.Value = 0
    Prg1.Min = 0
    Prg1.Max = Len(TxtSearch.Text)
    ShowProgress Prg1, StatusBar1, True
    For I = 0 To Len(TxtSearch.Text)
        x = TxtSearch.Find(sWord, I, , MatchCase)
        If x = -1 Then
            SearchFile = iFound: Me.MousePointer = 0: ShowProgress Prg1, StatusBar1, False: Exit Function
        Else
            iFound = iFound + 1: I = x
        End If
        Prg1.Value = I
    Next
    Me.MousePointer = 0
    SearchFile = iFound
    ShowProgress Prg1, StatusBar1, False
End Function

Private Sub SaveWalkthroughs()
    Dim SW
    SW = FreeFile
    If Dir(App.Path & "\Walkthroughs.ini") <> "" Then Kill App.Path & "\Walkthroughs.ini"
    If ListView1.ListItems.Count = 0 Then Exit Sub
    Open App.Path & "\Walkthroughs.ini" For Output As #SW
    For I = 1 To ListView1.ListItems.Count
        Print #SW, ListView1.ListItems(I) & "#" & ListView1.ListItems(I).SubItems(1) & "#" & ListView1.ListItems(I).SubItems(2) & "#" & lstWalkthru.List(I - 1)
    Next
    Close #SW
End Sub

Private Sub LoadWalkthroughs()
    On Error GoTo vbError
    Dim LW
    LW = FreeFile
    If Dir(App.Path & "\Walkthroughs.ini") = "" Then Exit Sub
    If FileLen(App.Path & "\Walkthroughs.ini") = 0 Then Exit Sub
    Open App.Path & "\Walkthroughs.ini" For Input As #LW
    Do While Not EOF(1)
        Input #LW, sData
        FileData() = Split(sData, "#")
        AddWalkthrough FileData(0), FileData(1), FileData(2), FileData(3)
    Loop
    Close #LW
vbError:
    Close #LW
    Exit Sub
End Sub

Private Sub mnuEdit_Click()
    Dim Platform, Genre, Title, FilePath, x
    If ListView1.ListItems.Count = 0 Then Exit Sub
    Platform = ListView1.ListItems(ListView1.SelectedItem.Index)
    Genre = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
    Title = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
    x = lstWalkthru.List(ListView1.SelectedItem.Index - 1)
    FilePath = GetFilePath(x, False)
    If FilePath <> "" Then
        Platform = InputBoxEx("What platform is this game for?", "Platform", Platform, , , , , , , "Courier New", "10")
        If Platform <> "" Then
            Genre = InputBoxEx("What is the genre of this game?", "Genre", Genre, , , , , , , "Courier New", "10")
            If Genre <> "" Then
                Title = InputBoxEx("What is the title of the game?", "Title", Title, , , , , , , "Courier New", "10")
                If Title <> "" Then
                    AddWalkthrough Platform, Genre, Title, FilePath, ListView1.SelectedItem.Index
                End If
            End If
        End If
    End If
End Sub

Private Sub mnuInsert_Click()
    Dim Platform, Genre, Title, FilePath
    FilePath = GetFilePath()
    If FilePath <> "" Then
        Platform = InputBoxEx("What platform is this game for?", "Platform", , , , , , , , "Courier New", "10")
        If Platform <> "" Then
            Genre = InputBoxEx("What is the genre of this game?", "Genre", , , , , , , , "Courier New", "10")
            If Genre <> "" Then
                Title = InputBoxEx("What is the title of the game?", "Title", , , , , , , , "Courier New", "10")
                If Title <> "" Then
                    AddWalkthrough Platform, Genre, Title, FilePath
                End If
            End If
        End If
    End If
End Sub

Private Sub DeleteWalkthrough(iIndex As Integer)
    If MsgBox("Remove " & Chr(34) & ListView1.ListItems(iIndex).SubItems(2) & Chr(34) & " from your walkthrough list ?", vbExclamation + vbYesNo, "Delete") = vbYes Then
        ListView1.ListItems.Remove iIndex
        lstWalkthru.RemoveItem iIndex - 1
        StatusBar1.Panels(1).Text = "Number of walkthroughs: " & ListView1.ListItems.Count
    End If
End Sub

Private Sub AddWalkthrough(sPlatform, sGenre, sTitle, sFilePath, Optional sEdit As Integer = -1)
    If sEdit = -1 Then
        lstWalkthru.AddItem sFilePath
        ListView1.ListItems.Add , , sPlatform
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = sGenre
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = sTitle
        If Dir(sFilePath) = "" Then ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "N/A": SetColour ListView1.ListItems.Count, vbRed, True Else ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = GetFileLength(sFilePath)
    Else
        lstWalkthru.List(sEdit - 1) = sFilePath
        ListView1.ListItems(sEdit).Text = sPlatform
        ListView1.ListItems(sEdit).SubItems(1) = sGenre
        ListView1.ListItems(sEdit).SubItems(2) = sTitle
        If Dir(sFilePath) = "" Then ListView1.ListItems(sEdit).SubItems(3) = "N/A": SetColour sEdit, vbRed, True Else ListView1.ListItems(sEdit).SubItems(3) = GetFileLength(sFilePath): SetColour sEdit, vbBlack, False
    End If
    StatusBar1.Panels(1).Text = "Number of walkthroughs: " & ListView1.ListItems.Count
End Sub

Private Function GetFileLength(sFileName)
    GetFileLength = Int(FileLen(sFileName) / 1024)
End Function

Private Sub SetColour(iIndex As Integer, lColour As Long, bBold As Boolean)
    With ListView1
        .ListItems(iIndex).ForeColor = lColour
        .ListItems(iIndex).Bold = bBold
        .ListItems(iIndex).ListSubItems(1).ForeColor = lColour
        .ListItems(iIndex).ListSubItems(1).Bold = bBold
        .ListItems(iIndex).ListSubItems(2).ForeColor = lColour
        .ListItems(iIndex).ListSubItems(2).Bold = bBold
        .ListItems(iIndex).ListSubItems(3).ForeColor = lColour
        .ListItems(iIndex).ListSubItems(3).Bold = bBold
    End With
End Sub

Private Function GetFilePath(Optional sFileName = "", Optional bShowOpen As Boolean = True) As String
    On Error GoTo vbError
    With CD
        .DialogTitle = "Walkthough location"
        .CancelError = True
        .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .FileName = sFileName
        .Flags = cdlOFNFileMustExist + cdlOFNPathMustExist + cdlOFNHideReadOnly
        If bShowOpen Then .ShowOpen Else .ShowSave
        If Len(.FileName) = 0 Then GoTo vbError
        GetFilePath = .FileName
    End With
    Exit Function
vbError:
    GetFilePath = ""
    Exit Function
End Function

Private Sub mnuOpen_Click()
    Dim y As Integer
    If ListView1.ListItems.Count <> 0 Then
        y = ListView1.SelectedItem.Index
        If ListView1.ListItems(y).ForeColor = vbRed Then
            x = lstWalkthru.List(y - 1)
            x = GetFilePath(x)
            If x <> "" Then AddWalkthrough ListView1.ListItems(y).Text, ListView1.ListItems(y).SubItems(1), ListView1.ListItems(y).SubItems(2), x, y: Shell "Notepad " & lstWalkthru.List(y - 1), vbNormalFocus
            RefreshList
        Else
            Shell "Notepad " & lstWalkthru.List(ListView1.SelectedItem.Index - 1), vbNormalFocus
        End If
    End If
End Sub

Private Sub mnuPrint_Click()
    If ListView1.ListItems.Count <> 0 And ListView1.ListItems(ListView1.SelectedItem.Index).ForeColor <> vbRed Then PrintFile lstWalkthru.List(ListView1.SelectedItem.Index - 1)
End Sub

Private Sub PrintFile(sFileName)
    On Error Resume Next
    TxtSearch.LoadFile sFileName
    With CD
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        .Flags = cdlPDHidePrintToFile + cdlPDNoSelection + cdlPDAllPages
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            TxtSearch.SelPrint .hdc
        End If
    End With
End Sub

Private Sub mnuRefresh_Click()
    RefreshList
End Sub

Private Sub mnuSearch_Click()
    Dim Word
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.ListItems(ListView1.SelectedItem.Index).ForeColor = vbRed Then MsgBox "Cannot search walkthrough because the file location is invalid.", vbCritical: Exit Sub
    Word = InputBoxEx("Word to search for:" & Chr(10) & "NB: End string with " & Chr(34) & " -mc" & Chr(34) & " to match case.", "Search", , , , , , , , "Courier New", "10")
    If Trim$(Word) <> "" Then MsgBox "The word " & Chr(34) & Word & Chr(34) & " was found " & SearchFile(lstWalkthru.List(ListView1.SelectedItem.Index - 1), Word) & " time(s).", vbInformation, "Search"
End Sub

Private Sub mnuRemAll_Click()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If MsgBox("Do you really want to remove all walkthroughs ?", vbExclamation + vbYesNo) = vbYes Then
        ListView1.ListItems.Clear
    End If
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Select Case Index
        Case 0
            mnuOpen_Click
        Case 1
            mnuPrint_Click
        Case 3
            mnuSearch_Click
        Case 5
            mnuInsert_Click
        Case 6
            mnuEdit_Click
        Case 7
            mnuDelete_Click
        Case 9
            mnuRefresh_Click
        Case 10
            mnuRemAll_Click
    End Select
End Sub
