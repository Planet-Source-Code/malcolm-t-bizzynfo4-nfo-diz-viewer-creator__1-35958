VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "bizzyNFO4 - About"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
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
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5400
      Top             =   4800
   End
   Begin VB.PictureBox pctHolder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   5625
      TabIndex        =   4
      Top             =   3600
      Width           =   5655
      Begin VB.Label lblInfoTxt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   $"frmAbout.frx":0000
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   960
         TabIndex        =   5
         Top             =   0
         Width           =   3735
      End
   End
   Begin VB.PictureBox pctBNFO4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      Picture         =   "frmAbout.frx":0130
      ScaleHeight     =   1785
      ScaleWidth      =   5865
      TabIndex        =   2
      Top             =   960
      Width           =   5895
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   5775
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Info"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   5895
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   2055
      Left            =   -120
      Top             =   840
      Width           =   6015
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1455
      Left            =   -120
      Top             =   3120
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   615
      Left            =   -120
      Top             =   4800
      Width           =   6015
   End
   Begin VB.Label lblAboutBizzyNFO4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About bizzyNFO4"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   -240
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
lblInfoTxt.Top = lblInfoTxt.Top - 10
    If lblInfoTxt.Top <= -2070 Then
        lblInfoTxt.Top = 0
    End If
End Sub
