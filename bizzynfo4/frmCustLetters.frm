VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCustLetters 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Objects"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4470
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
   ScaleHeight     =   5415
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   3720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4800
      Width           =   4215
   End
   Begin VB.CommandButton cmdAddSel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Insert Selected"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdAddToND 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Insert All"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3960
      Width           =   1935
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   35
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   34
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2400
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2400
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2400
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2400
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optTool 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin RichTextLib.RichTextBox txtCLetter 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmCustLetters.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HyperFont"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   615
      Left            =   -120
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Options"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   975
      Left            =   -120
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   2415
      Left            =   -240
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Custom Objects"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   -240
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu frmFile 
      Caption         =   "File"
      Begin VB.Menu frmSave 
         Caption         =   "Save.."
      End
      Begin VB.Menu frmOpen 
         Caption         =   "Open.."
      End
      Begin VB.Menu frmSapce 
         Caption         =   "-"
      End
      Begin VB.Menu frmCC 
         Caption         =   "Custom Characters.."
      End
      Begin VB.Menu frmSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu frmRTA 
         Caption         =   "Return To Application"
      End
   End
End
Attribute VB_Name = "frmCustLetters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddSel_Click()
frmMain.txtNFODIZ.SelText = frmMain.txtNFODIZ.SelText + txtCLetter.SelText
End Sub

Private Sub cmdAddToND_Click()
frmMain.txtNFODIZ.SelText = frmMain.txtNFODIZ.SelText + txtCLetter.Text
End Sub

Private Sub cmdReturn_Click()
On Error GoTo 1

ans = MsgBox("Do you want to save this bizzyNFO4 Custom Object file before you exit?", vbYesNo + vbQuestion, "bizzyNFO4")
    If ans = vbYes Then
        cd.DialogTitle = "Save a bizzyNFO4 Custom Object file.."
        cd.Filter = "bizzyNFO4 BCO File (*.bco)"
        cd.ShowSave
        txtCLetter.SaveFile cd.FileName
1     Unload Me
    ElseIf ans = vbNo Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    For i = 0 To 31
        optTool(i).Caption = frmTools.optChar(i).Caption
    Next i
End Sub

Private Sub frmCC_Click()
frmCCC2.Show
End Sub

Private Sub frmOpen_Click()
cd.DialogTitle = "Open a bizzyNFO4 Custom Object file.."
cd.Filter = "bizzyNFO4 BCO File (*.bco) | *.bco"
cd.ShowOpen

txtCLetter.LoadFile cd.FileName
End Sub

Private Sub frmRTA_Click()
On Error GoTo 11

ans = MsgBox("Do you want to save this bizzyNFO4 Custom Object file before you exit?", vbYesNo + vbQuestion, "bizzyNFO4")
    If ans = vbYes Then
        cd.DialogTitle = "Save a bizzyNFO4 Custom Object file.."
        cd.Filter = "bizzyNFO4 BCO File (*.bco)"
        cd.ShowSave
        txtCLetter.SaveFile cd.FileName
11     Unload Me
    ElseIf ans = vbNo Then
        Unload Me
    End If
    
End Sub

Private Sub frmSave_Click()
On Error Resume Next

cd.DialogTitle = "Save a bizzyNFO4 Custom Object file.."
cd.Filter = "bizzyNFO4 BCO File (*.bco)"
cd.ShowSave

txtCLetter.SaveFile cd.FileName
End Sub

Private Sub txtCLetter_KeyDown(KeyCode As Integer, Shift As Integer)
    For i = 0 To 35
       
       If optTool(i).Value = True Then
            
            If KeyCode = vbKeyF6 Then
                txtCLetter.SelText = txtCLetter.SelText + optTool(i).Caption
            End If
            
            Exit For
        End If
    
    Next i
End Sub
