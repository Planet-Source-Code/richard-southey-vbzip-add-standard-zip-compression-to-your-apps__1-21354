VERSION 5.00
Begin VB.Form frmDelete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmDelete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame fraFiles 
      Caption         =   "Files"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtFiles 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Text            =   "*.*"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optFiles 
         Caption         =   "&Files:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optSelected 
         Caption         =   "&Selected Files"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optEntire 
         Caption         =   "&Entire Archive"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmDelete"
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
    'Close the dialog
    Unload Me
End Sub


Private Sub cmdDelete_Click()
    Dim i As Long
    Dim Files As New Collection
    
    'Find which option was selected
    If optEntire.Value = True Then
        'Delete all the archive
        i = MsgBox("Are you sure you wish to delete " & frmArchive.VBZip.Filename, vbYesNo Or vbQuestion, "Delete Archive")
        If i = vbYes Then
            'Make sure the file still exists and selete it
            If Dir$(frmArchive.VBZip.Filename) <> "" Then
                Kill frmArchive.VBZip.Filename
            End If
            'Update VBZip
            frmArchive.VBZip.Filename = ""
        End If
        Unload Me
    End If
    
    If optSelected.Value = True Then
        'Delete the selected items
        'Make the collection
        For i = 1 To frmArchive.lvwArchive.ListItems.count
            With frmArchive.lvwArchive.ListItems(i)
                If .Selected Then
                    Files.Add frmArchive.VBZip.ParseFilename(frmArchive.VBZip.GetEntry(CLng(.Tag)).Filename)
                End If
            End With
        Next i
        'Close this dialog and delete the files
        Unload Me
        frmArchive.VBZip.Delete Files
    End If
    
    If optFiles.Value = True Then
        'Delete the files matching the wildcard entered
        Files.Add txtFiles.Text
        Unload Me
        frmArchive.VBZip.Delete Files
    End If
End Sub


