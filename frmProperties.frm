VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Properties"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdoK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblHexCRC 
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblFilePath 
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label lblDate 
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label lblCompRatio 
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblCompressed 
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblSize 
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblFilename 
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblCRC 
      Caption         =   "CRC:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblPath 
      Caption         =   "Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblModified 
      Caption         =   "Modified:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblRatio 
      Caption         =   "Ratio:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblCompSize 
      Caption         =   "Compressed size:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblFilesize 
      Caption         =   "File size:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmProperties"
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


Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'File in the details
    With frmArchive.VBZip.GetEntry(frmArchive.lvwArchive.SelectedItem.Tag)
        lblFilename.Caption = frmArchive.VBZip.ParseFilename(.Filename)
        lblSize.Caption = Format(.UncompressedSize, "###,###")
        lblCompressed.Caption = Format(.CompressedSize, "###,###")
        'Trap division by zero
        If .UncompressedSize <> 0 Then
            lblCompRatio.Caption = Format(CInt((1 - (.CompressedSize / .UncompressedSize)) * 100)) & "%"
        Else
            lblCompRatio.Caption = "0%"
        End If
        lblDate.Caption = .FileDateTime
        lblFilePath.Caption = frmArchive.VBZip.ParsePath(.Filename)
        lblHexCRC.Caption = Hex$(.CRC32)
    End With
End Sub


