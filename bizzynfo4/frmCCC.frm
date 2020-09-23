VERSION 5.00
Begin VB.Form frmCCC 
   BackColor       =   &H00400000&
   Caption         =   "bizzyNFO4 - Change Custom Characters"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3825
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
   ScaleHeight     =   5040
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4440
      Width           =   3495
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox txtCustom10 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   1800
      TabIndex        =   12
      Text            =   "C"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtCustom9 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   720
      TabIndex        =   11
      Text            =   "C"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtCustom8 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   2880
      TabIndex        =   10
      Text            =   "C"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtCustom7 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   1800
      TabIndex        =   9
      Text            =   "C"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtCustom6 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   720
      TabIndex        =   8
      Text            =   "C"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtCustom5 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "C"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtCustom4 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   720
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "C"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtCustom3 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "C"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtCustom2 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "C"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtCustom1 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   720
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "C"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1095
      Left            =   -720
      Top             =   3840
      Width           =   4215
   End
   Begin VB.Label lblNum10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#10:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1200
      TabIndex        =   22
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblNum9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#9:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblNum8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#8:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2280
      TabIndex        =   20
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblNum7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#7:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1200
      TabIndex        =   19
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblNum6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#6:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblNum5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#5:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1200
      TabIndex        =   17
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblNum4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#4:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblNum3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#3:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2280
      TabIndex        =   15
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblNum2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#2:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1200
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblNum1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#1:"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblC610 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Custom Characters 6 thru 10"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label lblC15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Custom Characters 1 thru 5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   2775
      Left            =   -960
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblCCC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change Custom Character"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   -240
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmCCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
ans = MsgBox("Are you sure you want to change the custom characters to these selected?", vbYesNo + vbQuestion, "bizzyNFO4")
    If ans = vbYes Then
        frmTools.optChar(85).Caption = txtCustom1.Text
        frmTools.optChar(86).Caption = txtCustom2.Text
        frmTools.optChar(87).Caption = txtCustom3.Text
        frmTools.optChar(88).Caption = txtCustom4.Text
        frmTools.optChar(89).Caption = txtCustom5.Text
        frmTools.optChar(90).Caption = txtCustom6.Text
        frmTools.optChar(91).Caption = txtCustom7.Text
        frmTools.optChar(92).Caption = txtCustom8.Text
        frmTools.optChar(93).Caption = txtCustom9.Text
        frmTools.optChar(94).Caption = txtCustom10.Text
        Unload Me
    ElseIf ans = vbNo Then
        'do nothing
    End If
End Sub

Private Sub Form_Load()
txtCustom1.Text = frmTools.optChar(85).Caption
txtCustom2.Text = frmTools.optChar(86).Caption
txtCustom3.Text = frmTools.optChar(87).Caption
txtCustom4.Text = frmTools.optChar(88).Caption
txtCustom5.Text = frmTools.optChar(89).Caption
txtCustom6.Text = frmTools.optChar(90).Caption
txtCustom7.Text = frmTools.optChar(91).Caption
txtCustom8.Text = frmTools.optChar(92).Caption
txtCustom9.Text = frmTools.optChar(93).Caption
txtCustom10.Text = frmTools.optChar(94).Caption
End Sub
