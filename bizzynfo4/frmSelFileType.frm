VERSION 5.00
Begin VB.Form frmSelFileType 
   BackColor       =   &H00400000&
   Caption         =   "Select A File Type.."
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
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
   ScaleHeight     =   2685
   ScaleWidth      =   2565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Continue.."
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.OptionButton optNFO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "NFO File"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.OptionButton optDIZ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DIZ File"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Shape shpBackGround3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   615
      Left            =   -360
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label lblSelFileType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select A File Type.."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Shape shpBackGround2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   855
      Left            =   -600
      Top             =   840
      Width           =   2895
   End
   Begin VB.Shape shpBackGround1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   -120
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmSelFileType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdContinue_Click()
    If optNFO.Value = True Then
        frmMDI.Show
        frmMDI.Caption = "bizzyNFO4 - Untitled.nfo - 0 Character(s)"
        frmInfo.lblFileType.Caption = "NFO"
        Unload Me
        Exit Sub
    End If
    
    If optDIZ.Value = True Then
        frmMDI.Show
        frmMDI.Caption = "bizzyNFO4 - Untitled.diz - 0 Character(s)"
        frmInfo.lblFileType.Caption = "DIZ"
        Unload Me
        Exit Sub
    End If
    
    If optDIZ.Value = False And optNFO.Value = False Then
        MsgBox "You must choose either DIZ or NFO before proceeding.", vbOKOnly + vbExclamation, "bizzyNFO4"
        Exit Sub
    End If
    
End Sub
