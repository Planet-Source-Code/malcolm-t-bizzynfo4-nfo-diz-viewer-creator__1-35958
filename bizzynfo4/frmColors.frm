VERSION 5.00
Begin VB.Form frmColors 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colors"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2565
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   2565
   Begin VB.Label lblFViolet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   22
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblFBlue 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblFTeal 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblFGreen 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblFYellow 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblFOrange 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblFRed 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblFBlack 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblBViolet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblBBlue 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblBTeal 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblBGreen 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblBYellow 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblBOrange 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblBRed 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblBWhite 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblRealPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This is a preview."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label lblPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Preview"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Shape shpBackGround4 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   855
      Left            =   0
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label lblFSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblBSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblFColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Foreground Color"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Shape shpBackGround3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   2055
      Left            =   0
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblBColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Background Color"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Shape shpBackGround2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   2055
      Left            =   -240
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Colors"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Shape shpBackGround1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   -360
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub lblBBlue_Click()
lblBSelected.BackColor = lblBBlue.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblBGreen_Click()
lblBSelected.BackColor = lblBGreen.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblBOrange_Click()
lblBSelected.BackColor = lblBOrange.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblBRed_Click()
lblBSelected.BackColor = lblBRed.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblBTeal_Click()
lblBSelected.BackColor = lblBTeal.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblBViolet_Click()
lblBSelected.BackColor = lblBViolet.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblBWhite_Click()
lblBSelected.BackColor = lblBWhite.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblBYellow_Click()
lblBSelected.BackColor = lblBYellow.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblfBlue_Click()
lblFSelected.BackColor = lblFBlue.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblfGreen_Click()
lblFSelected.BackColor = lblFGreen.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblfOrange_Click()
lblFSelected.BackColor = lblFOrange.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblfRed_Click()
lblFSelected.BackColor = lblFRed.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub
Private Sub lblfTeal_Click()
lblFSelected.BackColor = lblFTeal.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblfViolet_Click()
lblFSelected.BackColor = lblFViolet.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblfblack_Click()
lblFSelected.BackColor = lblFBlack.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub lblfYellow_Click()
lblFSelected.BackColor = lblFYellow.BackColor
lblRealPreview.BackColor = lblBSelected.BackColor
lblRealPreview.ForeColor = lblFSelected.BackColor
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelColor = lblFSelected.BackColor
frmMain.txtNFODIZ.BackColor = lblBSelected.BackColor
End Sub

Private Sub Timer1_Timer()
Label1.Caption = frmColors.Top
Label2.Caption = frmColors.Left
End Sub
