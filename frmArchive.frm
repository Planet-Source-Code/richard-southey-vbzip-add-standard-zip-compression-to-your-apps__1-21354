VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\AVBZip_Control.vbp"
Begin VB.Form frmArchive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Richsoft VBZip"
   ClientHeight    =   4680
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7335
   Icon            =   "frmArchive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VBZip_Control.RichsoftVBZip VBZip 
      Left            =   5280
      Top             =   3720
      _ExtentX        =   1720
      _ExtentY        =   1720
   End
   Begin MSComDlg.CommonDialog cdlZip 
      Left            =   4680
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Richsoft VBZip"
      Filter          =   "Zip Files|*.zip"
   End
   Begin ComctlLib.ListView lvwArchive 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "iglIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date/Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Packed"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ratio"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblFiles 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Label lblRichsoft 
      BackStyle       =   0  'Transparent
      Caption         =   "Richsoft Computing 2001"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label lblWeb 
      Caption         =   "www.richsoftcomputing.btinternet.co.uk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmArchive.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Tag             =   "www.richsoftcomputing.btinternet.co.uk"
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "richsoftcomputing@btinternet.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmArchive.frx":045C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Tag             =   "mailto:richsoftcomputing@btinternet.com?subject=Richsoft VBZip"
      Top             =   4080
      Width           =   3135
   End
   Begin ComctlLib.ImageList iglIcons 
      Left            =   6480
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmArchive.frx":05AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuActionS 
      Caption         =   "&Actions"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add..."
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete..."
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "&Extract..."
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View..."
      End
      Begin VB.Menu mnuActionSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuInvert 
         Caption         =   "Invert Selection"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Richsoft VBZip..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupExtract 
         Caption         =   "&Extract..."
      End
      Begin VB.Menu mnuPopupDelete 
         Caption         =   "&Delete..."
      End
      Begin VB.Menu mnuPopupView 
         Caption         =   "&View..."
      End
      Begin VB.Menu mnuPopupSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuPopupInvert 
         Caption         =   "Invert Selection"
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupProperties 
         Caption         =   "&File Properties..."
      End
   End
End
Attribute VB_Name = "frmArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'==============================================================================
'Richsoft Computing 2001
'Richard Southey
'This code is e-mailware, if you use it please e-mail me and tell me about
'your program.
'
'For latest information about this and other projects please visit my website:
'www.richsoftcomputing.btinternet.co.uk
'
'If you would like to make any comments/suggestions then please e-mail them to
'richsoftcomputing@btinternet.co.uk
'==============================================================================

'API Call which drives the Hyperlink
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Sub HyperJump(ByVal URL As String)
    'Function to execute the Hyperlink
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub


Private Sub lblEmail_Click()
    'Send an email
    HyperJump lblEmail.Tag
End Sub

Private Sub lblWeb_Click()
    'Go to my website
    HyperJump lblWeb.Tag
End Sub


Private Sub lvwArchive_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    'Sort by the column clicked
    lvwArchive.Sorted = True
    lvwArchive.SortKey = ColumnHeader.Index - 1
    
End Sub


Private Sub lvwArchive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    Dim count As Long

    If lvwArchive.ListItems.count = 0 Then Exit Sub
    
    'Check if the right button was pressed
    If Button <> vbRightButton Then Exit Sub
    
    'Check an item has been clicked on
    If lvwArchive.HitTest(x, y) Is Nothing Then Exit Sub
    
    'Check to see if the item under the mouse is selected
    If lvwArchive.HitTest(x, y).Selected Then
        'It's already selected so just show the popup menu
        PopupMenu mnuPopup
    Else
        'Deselect other items and just select this one
        For i = 1 To lvwArchive.ListItems.count
            lvwArchive.ListItems(i).Selected = False
        Next i
        lvwArchive.HitTest(x, y).Selected = True
        PopupMenu mnuPopup
    End If
End Sub


Private Sub mnuAbout_Click()
    'Show the VBZip about box
    VBZip.About
End Sub

Private Sub mnuAdd_Click()
    'Show the add dialog
    If VBZip.Filename <> "" Then frmAdd.Show 1
End Sub


Private Sub mnuDelete_Click()
    'Show the delete dialog
    If VBZip.Filename <> "" Then frmDelete.Show 1
End Sub

Private Sub mnuExit_Click()
    'Show about box and exit
    VBZip.About
    Unload Me
End Sub

Private Sub mnuExtract_Click()
    'Show the extract dialog
    If VBZip.Filename <> "" Then
        frmExtract.Show 1
    End If
End Sub

Private Sub mnuInvert_Click()
    'Inverts the current selection
    Dim i As Long
    For i = 1 To lvwArchive.ListItems.count
        lvwArchive.ListItems(i).Selected = Not (lvwArchive.ListItems(i).Selected)
    Next i
End Sub

Private Sub mnuOpen_Click()
    'Open an archive
    'Using a filename that does not exist will create a new archive
    Dim r As Integer
    
    On Error Resume Next
    cdlZip.ShowOpen
    'Check if cancel was pressed
    If Err = cdlCancel Then Exit Sub
    
    'Ask if the archive should be created if it does not exist
    If Dir$(cdlZip.Filename) = "" Then
        r = MsgBox("Do you wish to create " + VBZip.ParseFilename(cdlZip.Filename) + "?", _
            vbQuestion Or vbYesNo, "Create Archive")
        If r = vbNo Then Exit Sub
    End If
    
    VBZip.Filename = cdlZip.Filename
End Sub

Private Sub mnuPopupDelete_Click()
    'Delete files from the archive
    mnuDelete_Click
End Sub

Private Sub mnuPopupExtract_Click()
    'Extract files
    mnuExtract_Click
End Sub

Private Sub mnuPopupInvert_Click()
    'Invert the current selection
    Call mnuInvert_Click
End Sub

Private Sub mnuPopupProperties_Click()
    'Show the property box
    frmProperties.Show 1, Me
End Sub

Private Sub mnuPopupSelectAll_Click()
    'Select all
    Call mnuSelectAll_Click
End Sub


Private Sub mnuPopupView_Click()
    'View all the selected items
    Call mnuView_Click
End Sub

Private Sub mnuSelectAll_Click()
    'Selects all the items
    Dim i As Long
    For i = 1 To lvwArchive.ListItems.count
        lvwArchive.ListItems(i).Selected = True
    Next i
End Sub


Private Sub mnuView_Click()
    'Extracts the file to TEMP and then opens it
    Dim i As Long
    Dim r As Long
    Dim Files As New Collection
    
    If lvwArchive.SelectedItem Is Nothing Then Exit Sub
    
    'Find all the selected items
    For i = 1 To lvwArchive.ListItems.count
        With lvwArchive.ListItems(i)
            If .Selected Then
                Files.Add VBZip.ParseFilename(VBZip.GetEntry(CLng(.Tag)).Filename)
            End If
        End With
    Next i
    
    'Extract the files to TEMP
    r = VBZip.Extract(Files, zipDefault, False, True, Environ$("TEMP"))
    'Open them if they extracted
    For i = 1 To Files.count
        If Dir$(Environ$("TEMP") + "\" + VBZip.ParseFilename(Files(i))) <> "" Then
            HyperJump Environ$("TEMP") + "\" + VBZip.ParseFilename(Files(i))
        End If
    Next i
    
End Sub

Private Sub VBZip_OnArchiveUpdate()
    'Fill the listview with the archives contents
    Dim itmX As ListItem
    Dim i As Long
    Dim Entry As ZipFileEntry
    Dim TotalSize As Long
    
    'Update the form caption
    Me.Caption = "Richsoft VBZip - " & VBZip.Filename
    
    'Clear the list
    lvwArchive.ListItems.Clear
    
    'Loop thought the entries, updating the listview
    'missing any blank entries
    With VBZip
        For i = 1 To .GetEntryNum
            Set Entry = .GetEntry(i)
            If Entry.Filename <> "" Then
                Set itmX = lvwArchive.ListItems.Add(, , .ParseFilename(Entry.Filename), , 1)
                itmX.SubItems(1) = Entry.FileDateTime
                itmX.SubItems(2) = Format(Entry.UncompressedSize, "###,###")
                itmX.SubItems(3) = Format(Entry.CompressedSize, "###,###")
                'Trap division by zero
                If Entry.UncompressedSize <> 0 Then
                    itmX.SubItems(4) = Format(CInt((1 - (Entry.CompressedSize / Entry.UncompressedSize)) * 100)) & "%"
                Else
                    itmX.SubItems(4) = "0%"
                End If
                itmX.SubItems(5) = .ParsePath(Entry.Filename)
                'Save the item number for other operations
                itmX.Tag = i
                TotalSize = TotalSize + Entry.UncompressedSize
            End If
        Next i
        
        lblFiles.Caption = CStr(lvwArchive.ListItems.count) + " file(s), " + VBZip.ConvertBytesToString(TotalSize)
    End With
End Sub



Private Sub VBZip_OnDeleteComplete(ByVal Successful As Long)
    Unload frmProgress
End Sub


Private Sub VBZip_OnDeleteProgress(ByVal Percentage As Integer, ByVal Filename As String)
    With frmProgress
        .Show , Me
        .pbrProgress.Value = Percentage
        .lblWorking = "Deleting " + Filename + "..."
    End With
End Sub


Private Sub VBZip_OnUnzipComplete(ByVal Successful As Long)
    Unload frmProgress
End Sub

Private Sub VBZip_OnUnzipProgress(ByVal Percentage As Integer, ByVal Filename As String)
    With frmProgress
        .Show , Me
        .pbrProgress.Value = Percentage
        .lblWorking = "Extracting " + Filename + "..."
    End With
End Sub


Private Sub VBZip_OnZipComplete(ByVal Successful As Long)
    Unload frmProgress
End Sub

Private Sub VBZip_OnZipProgress(ByVal Percentage As Integer, ByVal Filename As String)
    With frmProgress
        .Show , Me
        .pbrProgress.Value = Percentage
        .lblWorking = "Adding " + Filename + "..."
    End With
End Sub


