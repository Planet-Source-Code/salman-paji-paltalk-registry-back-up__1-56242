VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PalTalkRegistryBackUp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PalTalk Registry BackUp VB Project by SP"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      Height          =   300
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7200
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.reg"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Choose Your Saving Path For The Back-up by Selecting The ""Save"" Button"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "PalTalkRegistryBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PalTalkReg(Path As String)

Shell "Regedit.exe /e " & Chr(34) & Path & Chr(34) & " " & Chr(34) & _
"HKEY_CURRENT_USER" & "\" & "Software\PalTalk\" & Chr(34)
'The Char For " is Chr(34)

End Sub

Private Sub Save_Click()

On Error Resume Next

CD.DialogTitle = "Save PalTalk Registry Backup File As..."
CD.InitDir = App.Path
CD.Flags = &H4
CD.Filter = "Registry File Format (*.reg)|*.reg"
CD.ShowSave

If InStr(CD.FileName, ".reg") Then

Call PalTalkReg(CD.FileName)

DoEvents

Text1.Text = CD.FileName

End If

End Sub
