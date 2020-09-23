VERSION 5.00
Begin VB.Form frmTChooser 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Tool"
   ClientHeight    =   5460
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
   ScaleHeight     =   5460
   ScaleWidth      =   2565
   Begin VB.OptionButton optSel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optSel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optSel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optSel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optSel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optSel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optSel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblSel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1200
      TabIndex        =   23
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblSel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   22
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblSel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   21
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblSel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   20
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblSel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   19
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblSel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   18
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblSel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   17
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblTool7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tool #7:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lblTool6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tool #6:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lblTool5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tool #5:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label lblTool4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tool #4:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblTool3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tool #3:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblTool2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tool #2:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblTool1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tool #1:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblCharForTool 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Character For Tool #"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   3015
      Left            =   -240
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label lblTool 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tool Number"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblSTool 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Set Tool"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1215
      Left            =   -240
      Top             =   840
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   -120
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmTChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub optSel_Click(Index As Integer)
    For i = 0 To 6
        If optSel(i).Value = True Then
            lblSel(i).Caption = frmInfo.lblSelectedTool
            Exit For
        End If
    Next i
End Sub
