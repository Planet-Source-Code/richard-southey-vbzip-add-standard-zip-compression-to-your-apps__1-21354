VERSION 5.00
Begin VB.Form frmExtract 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extract"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.DriveListBox drvListing 
      Height          =   315
      Left            =   3600
      TabIndex        =   9
      Top             =   3000
      Width           =   2535
   End
   Begin VB.DirListBox dirListing 
      Height          =   2340
      Left            =   3600
      TabIndex        =   7
      Top             =   480
      Width           =   2535
   End
   Begin VB.CheckBox chkFolders 
      Caption         =   "&Use folder names"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CheckBox chkSkipOlder 
      Caption         =   "S&kip older files"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CheckBox chkOverwrite 
      Caption         =   "&Overwrite existing files"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Frame fraFiles 
      Caption         =   "Files"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3375
      Begin VB.TextBox txtFiles 
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Text            =   "*.*"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optFiles 
         Caption         =   "&Files:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optAllFiles 
         Caption         =   "&All Files"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton optSelected 
         Caption         =   "&Selected Files"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdShowPath 
      Caption         =   "->"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtExtractTo 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblFolders 
      Caption         =   "Folders/drives:"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblExtractTo 
      Caption         =   "Extract to:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmExtract"
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
Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdExtract_Click()
    Dim Files As New Collection
    Dim Path As String
    Dim r As Long
    'Extract the files
    If optSelected.Value = True Then
        'Find all the selected items and add them to the collection
        For i = 1 To frmArchive.lvwArchive.ListItems.count
            With frmArchive.lvwArchive.ListItems(i)
                If .Selected Then
                    Files.Add frmArchive.VBZip.ParseFilename(frmArchive.VBZip.GetEntry(CLng(.Tag)).Filename)
                End If
            End With
        Next i
    End If
    
    If optAllFiles.Value = True Then
        'Add all files
        For i = 1 To frmArchive.lvwArchive.ListItems.count
            With frmArchive.lvwArchive.ListItems(i)
                Files.Add frmArchive.VBZip.ParseFilename(frmArchive.VBZip.GetEntry(CLng(.Tag)).Filename)
            End With
        Next i
    End If
    
    If optFiles.Value = True Then
        Files.Add txtFiles.Text
    End If
    
    'Check extract dir exists
    If Dir$(txtExtractTo.Text, vbDirectory) = "" Then
        MsgBox txtExtractTo.Text & " does not exist!", vbCritical, "Extract folder does not exist"
        Exit Sub
    End If
    
    Path = txtExtractTo.Text
    
    Unload Me
    'Now extract all the files
    'chkSkipOlder will be 1 for zipFreshen, 0 for zipDefault
    
    r = frmArchive.VBZip.Extract(Files, chkSkipOlder.Value, CBool(-chkFolders.Value), CBool(-chkOverwrite), Path)
End Sub


Private Sub cmdShowPath_Click()
    'Try and show the path entered in the folder list
    On Error GoTo errorHandler
    
    dirListing.Path = txtExtractTo.Text
    drvListing.Drive = dirListing.Path
    Exit Sub
    
errorHandler:
    'Could not find the folder
    MsgBox "Could not find folder " & txtExtractTo.Text, vbCritical, "Richsoft VBZip"
    Exit Sub
End Sub




Private Sub dirListing_Change()
    'Update the test box
    txtExtractTo.Text = dirListing.Path
End Sub

Private Sub drvListing_Change()
    'Change the directory lsiting to the current directory
    'of the select drive
    'Trap errors
    On Error GoTo errorHandler
    
    dirListing.Path = CurDir$(drvListing.Drive)
    
    Exit Sub
    
errorHandler:
    'The error handler
    MsgBox "The drive selected cannot be accessed", vbCritical, APP_TITLE
    Exit Sub
    Resume Next

End Sub


