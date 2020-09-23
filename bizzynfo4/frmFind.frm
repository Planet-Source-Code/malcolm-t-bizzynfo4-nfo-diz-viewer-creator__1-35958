VERSION 5.00
Begin VB.Form frmFind 
   BackColor       =   &H00400000&
   Caption         =   "bizzyNFO4 - Find"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFindNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find Next"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   3735
   End
   Begin VB.CommandButton cmdReplace 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Replace Selected"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   3735
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find 1st"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtReplace 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1080
      TabIndex        =   6
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lblCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Return To Application"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   3975
   End
   Begin VB.Shape shpBackground4 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   975
      Left            =   -120
      Top             =   4200
      Width           =   4095
   End
   Begin VB.Label lblReplace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Replace:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblReplaceWith 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Replace With"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Shape shpBackground3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1455
      Left            =   -120
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label lblWhatToFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "What To Find"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lblFindd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Find:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Shape shpBackground2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1455
      Left            =   -120
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label lblFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Find"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Shape shpBackground1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   -480
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
frmMain.txtNFODIZ.SetFocus
frmMain.txtNFODIZ.Find txtFind, 0, Len(frmMain.txtNFODIZ.Text)
End Sub

Private Sub cmdReplace_Click()
frmMain.txtNFODIZ.SelText = txtReplace
End Sub

Private Sub CmdFindNext_Click()
frmMain.txtNFODIZ.SetFocus
frmMain.txtNFODIZ.Find txtFind, frmMain.txtNFODIZ.SelStart + Len(frmMain.txtNFODIZ.SelText), Len(frmMain.txtNFODIZ.Text)
End Sub
