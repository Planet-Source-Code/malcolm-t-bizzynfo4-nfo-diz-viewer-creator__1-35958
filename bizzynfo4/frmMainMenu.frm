VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "bizzyNFO4 - Main Menu"
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
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
   ScaleHeight     =   4110
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3840
      Top             =   2400
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2760
      Picture         =   "frmMainMenu.frx":0000
      ScaleHeight     =   825
      ScaleWidth      =   2145
      TabIndex        =   0
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2002 BizZy"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label lblExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit bizzyNFO Version 4.0."
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblLoad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Load An Existing NFO / DIZ File."
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblCreate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Create A New NFO / DIZ File."
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMainMenu.frx":1403
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Shape shpBackGround4 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1215
      Left            =   2880
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Shape shpBackGround3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   2775
      Left            =   2880
      Top             =   0
      Width           =   2535
   End
   Begin VB.Shape shpBackGround2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   3255
      Left            =   -360
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Shape shpBackGround1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1695
      Left            =   -600
      Top             =   -840
      Width           =   3015
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TitleOfND As String
Dim LengthOfND As String
Private Sub Form_Load()
lblTime = Time
lblDate = Date
End Sub

Private Sub lblCreate_Click()
frmSelFileType.Show
Unload Me
End Sub

Private Sub lblCreate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCreate.BackColor = &H800000
lblCreate.ForeColor = vbWhite
lblDescription.Caption = "In the mood to create a new NFO or DIZ file? Well click here to begin making that NFO or DIZ file!"
    If lblExit.ForeColor = vbWhite Or lblLoad.ForeColor = vbWhite Then
    lblExit.BackColor = vbWhite
    lblExit.ForeColor = vbBlack
    lblLoad.BackColor = vbWhite
    lblLoad.ForeColor = vbBlack
    End If
End Sub

Private Sub lblExit_Click()
ans = MsgBox("Are you sure you want to exit bizzyNFO4?", vbYesNo + vbInformation, "bizzyNFO4")
    If ans = vbYes Then
        End
    ElseIf ans = vbNo Then
        Exit Sub
    End If
End Sub

Private Sub lblLoad_Click()
On Error GoTo 12
frmMDI.Show
frmMDI.Caption = "bizzyNFO4 - Untitled.nfo - 0 Character(s)"
frmMDI.cd.DialogTitle = "Load NFO \ DIZ File.."
frmMDI.cd.Filter = "bizzyNFO4 NFO File (*.nfo) |*.nfo|bizzyNFO4 DIZ File (*.diz)|*.diz"
frmMDI.cd.ShowOpen

frmMain.txtNFODIZ.LoadFile frmMDI.cd.FileName
TitleOfND = frmMDI.cd.FileTitle
LengthOfND = Len(frmMain.txtNFODIZ.Text)
frmMDI.Caption = "bizzyNFO4 - " + TitleOfND + " - " + LengthOfND + " Character(s)"

    If LCase(Mid(frmMDI.cd.FileTitle, Len(frmMDI.cd.FileTitle) - 2, Len(frmMDI.cd.FileTitle))) = "nfo" Then
        frmInfo.lblFileType.Caption = "NFO"
        Exit Sub
    End If
    
    If LCase(Mid(frmMDI.cd.FileTitle, Len(frmMDI.cd.FileTitle) - 2, Len(frmMDI.cd.FileTitle))) = "diz" Then
        frmInfo.lblFileType.Caption = "DIZ"
        Exit Sub
    End If
    
12    frmMDI.Caption = "bizzyNFO4 - Untitled.nfo - 0 Character(s)"
 
End Sub

Private Sub lblload_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLoad.ForeColor = vbWhite
lblLoad.BackColor = &H800000
lblDescription.Caption = "Have a NFO or DIZ file you were working on? Want to resume work on it? Well click here."
    If lblExit.ForeColor = vbWhite Or lblCreate.ForeColor = vbWhite Then
    lblExit.BackColor = vbWhite
    lblExit.ForeColor = vbBlack
    lblCreate.BackColor = vbWhite
    lblCreate.ForeColor = vbBlack
    End If
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbWhite
lblExit.BackColor = &H800000
lblDescription.Caption = "Exit the application."
    If lblLoad.ForeColor = vbWhite Or lblCreate.ForeColor = vbWhite Then
    lblLoad.BackColor = vbWhite
    lblLoad.ForeColor = vbBlack
    lblCreate.BackColor = vbWhite
    lblCreate.ForeColor = vbBlack
    End If
End Sub

Private Sub Timer1_Timer()
lblTime = Time
lblDate = Date
End Sub
