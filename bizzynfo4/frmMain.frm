VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00400000&
   Caption         =   "bizzyNFO4 Workspace"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4935
   Begin RichTextLib.RichTextBox txtNFODIZ 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6588
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":0ECA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HyperFont"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TurnedOn 
      Caption         =   "False"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TitleOfND2 As String
Dim TitleOfND As String
Dim LengthOfND As String
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
End Sub

Private Sub Form_Resize()
txtNFODIZ.Width = Val(frmMain.Width) - 120
txtNFODIZ.Height = Val(frmMain.Height) - 405
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub txtNFODIZ_Change()
    If frmMDI.cd.FileTitle <> "" Then
        TitleOfND = frmMDI.cd.FileTitle
        LengthOfND = Len(txtNFODIZ.Text)
        frmMDI.Caption = "bizzyNFO4 - " + TitleOfND + " - " + LengthOfND + " Character(s)"
    End If
    
    If frmMDI.cd.FileTitle = "" Then
        LengthOfND = Len(txtNFODIZ.Text)
        TitleOfND2 = LCase(frmInfo.lblFileType.Caption)
        TitleOfND = "Untitled." + TitleOfND2 + ""
        frmMDI.Caption = "bizzyNFO4 - " + TitleOfND + " - " + LengthOfND + " Character(s)"
    End If
End Sub

Private Sub txtNFODIZ_dblClick()
    If TurnedOn = True Then
        txtNFODIZ.SelText = txtNFODIZ.SelText + frmInfo.lblSelectedTool
    ElseIf TurnedOn = False Then
        MsgBox "Double Click ToolAdd must be on to do this. To turn it on goto Options/Double Click ToolAdd On.", vbOKOnly + vbInformation, "bizzyNFO4"
    End If
End Sub

Private Sub txtNFODIZ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
    txtNFODIZ.SelText = txtNFODIZ.SelText + frmInfo.lblSelectedTool
    End If
    
    If KeyCode = vbKeyF1 Then
    txtNFODIZ.SelText = txtNFODIZ.SelText + frmTChooser.lblSel(0).Caption
    End If
    
    If KeyCode = vbKeyF2 Then
    txtNFODIZ.SelText = txtNFODIZ.SelText + frmTChooser.lblSel(1).Caption
    End If
    
    If KeyCode = vbKeyF3 Then
    txtNFODIZ.SelText = txtNFODIZ.SelText + frmTChooser.lblSel(2).Caption
    End If
    
    If KeyCode = vbKeyF4 Then
    txtNFODIZ.SelText = txtNFODIZ.SelText + frmTChooser.lblSel(3).Caption
    End If
    
    If KeyCode = vbKeyF5 Then
    txtNFODIZ.SelText = txtNFODIZ.SelText + frmTChooser.lblSel(4).Caption
    End If
    
    If KeyCode = vbKeyF7 Then
    txtNFODIZ.SelText = txtNFODIZ.SelText + frmTChooser.lblSel(5).Caption
    End If
    
    If KeyCode = vbKeyF8 Then
    txtNFODIZ.SelText = txtNFODIZ.SelText + frmTChooser.lblSel(6).Caption
    End If
End Sub
