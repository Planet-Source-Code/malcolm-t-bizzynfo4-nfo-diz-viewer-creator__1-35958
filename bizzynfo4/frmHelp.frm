VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "bizzyNFO4 - Help"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
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
   ScaleHeight     =   8145
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picStep42 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4320
      Picture         =   "frmHelp.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   1215
      TabIndex        =   12
      Top             =   6120
      Width           =   1215
   End
   Begin VB.PictureBox pictStep41 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   600
      Picture         =   "frmHelp.frx":0B16
      ScaleHeight     =   855
      ScaleWidth      =   1215
      TabIndex        =   11
      Top             =   6120
      Width           =   1215
   End
   Begin VB.PictureBox pictStep3 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   3960
      Picture         =   "frmHelp.frx":105E
      ScaleHeight     =   975
      ScaleWidth      =   1575
      TabIndex        =   10
      Top             =   4800
      Width           =   1575
   End
   Begin VB.PictureBox pictStep2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   600
      Picture         =   "frmHelp.frx":19FE
      ScaleHeight     =   975
      ScaleWidth      =   1815
      TabIndex        =   9
      Top             =   3480
      Width           =   1815
   End
   Begin VB.PictureBox pctStep1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   4320
      Picture         =   "frmHelp.frx":2402
      ScaleHeight     =   975
      ScaleWidth      =   1215
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
      Width           =   5775
   End
   Begin VB.Label lblStep4txt 
      BackStyle       =   0  'Transparent
      Caption         =   "Explore other menu's and windows. They can do lots of things!"
      Height          =   855
      Left            =   1920
      TabIndex        =   16
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label lblStep3txt 
      BackStyle       =   0  'Transparent
      Caption         =   "To add a selected tool press F6. This will insert that tool into your NFO/DIZ File. Hold it down and you can insert more."
      Height          =   975
      Left            =   600
      TabIndex        =   15
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label lblStep2txt 
      BackStyle       =   0  'Transparent
      Caption         =   "Next click on the Workspace window. Here you can add a selected tool to your NFO/DIZ file."
      Height          =   975
      Left            =   2520
      TabIndex        =   14
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label lblStep1txt 
      BackStyle       =   0  'Transparent
      Caption         =   "First you must select a tool you want to work. To do this, select the tools windows, find the word characters, and choose a tool."
      Height          =   975
      Left            =   600
      TabIndex        =   13
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   615
      Left            =   -240
      Top             =   7440
      Width           =   6015
   End
   Begin VB.Label lblStep4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#4:"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   5535
   End
   Begin VB.Label lblStep3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#3:"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   5535
   End
   Begin VB.Label lblStep2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#2:"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   5535
   End
   Begin VB.Label lblStep1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#1:"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   5535
   End
   Begin VB.Label lblHelptxt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Here is some basic help to get you started on making your very own NFO/DIZ file."
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label lblSteps 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Steps"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   5775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   5655
      Left            =   -240
      Top             =   1560
      Width           =   6015
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Help"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1215
      Left            =   -840
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturn_Click()
Unload Me
End Sub

