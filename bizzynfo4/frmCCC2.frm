VERSION 5.00
Begin VB.Form frmCCC2 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "bizzyNFO4 - Change Custom Characters"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
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
   ScaleHeight     =   3855
   ScaleWidth      =   3405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   3135
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox txtChar4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtChar3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtChar2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtChar1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1095
      Left            =   -480
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label lblNum4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#4:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblNum3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#3:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblNum2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#2:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblCustom14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Custom Characters 1 Thru 4"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblNum1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#1:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1575
      Left            =   0
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblChange 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change Custom Characters"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   -360
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmCCC2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
ans = MsgBox("Are you sure you want to change the tools to these selected?", vbInformation + vbYesNo, "bizzyNFO4")
    If ans = vbYes Then
        frmCustLetters.optTool(32).Caption = txtChar1.Text
        frmCustLetters.optTool(33).Caption = txtChar2.Text
        frmCustLetters.optTool(34).Caption = txtChar3.Text
        frmCustLetters.optTool(35).Caption = txtChar4.Text
    ElseIf ans = vbNo Then
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
txtChar1.Text = frmCustLetters.optTool(32).Caption
txtChar3.Text = frmCustLetters.optTool(33).Caption
txtChar4.Text = frmCustLetters.optTool(34).Caption
txtChar5.Text = frmCustLetters.optTool(35).Caption
End Sub
