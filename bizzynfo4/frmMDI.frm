VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00400000&
   Caption         =   "bizzyNFO4 - Untitled.nfo - 0 Character(s)"
   ClientHeight    =   6330
   ClientLeft      =   2520
   ClientTop       =   1470
   ClientWidth     =   7305
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd 
      Left            =   120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu frmFile 
      Caption         =   "File"
      Begin VB.Menu frmNNFO 
         Caption         =   "New NFO.."
      End
      Begin VB.Menu frmLNFO 
         Caption         =   "Load NFO.."
      End
      Begin VB.Menu frmSNFO 
         Caption         =   "Save NFO.."
      End
      Begin VB.Menu frmSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu frmNDIZ 
         Caption         =   "New DIZ.."
      End
      Begin VB.Menu frmLDIZ 
         Caption         =   "Load DIZ.."
      End
      Begin VB.Menu frmSDIZ 
         Caption         =   "Save DIZ.."
      End
      Begin VB.Menu frmSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu frmPNFODIZ 
         Caption         =   "Print NFO \ DIZ.."
      End
      Begin VB.Menu frmSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu frmExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu frmEdit 
      Caption         =   "Edit"
      Begin VB.Menu frmESelAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu frmESelNone 
         Caption         =   "Select None"
      End
      Begin VB.Menu frmSpace4 
         Caption         =   "-"
      End
      Begin VB.Menu frmEFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu frmSpace5 
         Caption         =   "-"
      End
      Begin VB.Menu frmEcut 
         Caption         =   "Cut"
      End
      Begin VB.Menu frmECopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu frmEPaste 
         Caption         =   "Paste"
      End
   End
   Begin VB.Menu frmWindow 
      Caption         =   "Window"
      Begin VB.Menu frmColors1 
         Caption         =   "Colors"
         Checked         =   -1  'True
      End
      Begin VB.Menu frmInfo1 
         Caption         =   "Info"
         Checked         =   -1  'True
      End
      Begin VB.Menu frmMain1 
         Caption         =   "Workspace"
         Checked         =   -1  'True
      End
      Begin VB.Menu frmTools1 
         Caption         =   "Tools"
         Checked         =   -1  'True
      End
      Begin VB.Menu frmTChooser1 
         Caption         =   "Set Tool"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu frmAlignment 
      Caption         =   "Alignment"
      Begin VB.Menu frmALeft 
         Caption         =   "Left"
      End
      Begin VB.Menu frmACenter 
         Caption         =   "Center"
      End
      Begin VB.Menu frmARight 
         Caption         =   "Right"
      End
   End
   Begin VB.Menu frmOptions 
      Caption         =   "Options"
      Begin VB.Menu frmCCC1 
         Caption         =   "Change Custom Characters.."
      End
      Begin VB.Menu frmCF 
         Caption         =   "Change Font.."
      End
      Begin VB.Menu frmSpace6 
         Caption         =   "-"
      End
      Begin VB.Menu frmOTDCLTO 
         Caption         =   "Turn Double Click ToolAdd On"
      End
   End
   Begin VB.Menu frmCLetters1 
      Caption         =   "Custom Objects"
      Begin VB.Menu frmCLCustom 
         Caption         =   "Custom Objects Editor"
      End
   End
   Begin VB.Menu frmHelp1 
      Caption         =   "Help"
      Begin VB.Menu frmHHelp 
         Caption         =   "Basic Help.."
      End
      Begin VB.Menu frmHAbout 
         Caption         =   "About.."
      End
      Begin VB.Menu frmSpace7 
         Caption         =   "-"
      End
      Begin VB.Menu frmHTC 
         Caption         =   "Help With Tool Chooser"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LengthOfND As String
Dim TitleOfND As String


Private Sub frmACenter_Click()
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelAlignment = rtfCenter
frmMain.txtNFODIZ.SetFocus
End Sub

Private Sub frmALeft_Click()
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelAlignment = rtfLeft
frmMain.txtNFODIZ.SetFocus
End Sub

Private Sub frmARight_Click()
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
frmMain.txtNFODIZ.SelAlignment = rtfRight
frmMain.txtNFODIZ.SetFocus
End Sub

Private Sub frmCCC1_Click()
frmCCC.Show
End Sub

Private Sub frmCF_Click()
cd.Flags = cdlCFBoth Or cdlCFForceFontExist Or cdlCFEffects
cd.FontName = frmMain.txtNFODIZ.Font.Name
cd.ShowFont


frmMain.txtNFODIZ.Font.Name = cd.FontName
frmMain.txtNFODIZ.Font.Size = cd.FontSize
frmMain.txtNFODIZ.Font.Strikethrough = cd.FontStrikethru
frmMain.txtNFODIZ.Font.Underline = cd.FontUnderline
1 End Sub

Private Sub frmCLetters_Click()

End Sub

Private Sub frmCLCustom_Click()
frmCustLetters.Show
End Sub

Private Sub frmColors1_Click()

    If frmColors1.Checked = True Then
        frmColors.Hide
        frmColors1.Checked = False
        Exit Sub
    End If
    
    If frmColors1.Checked = False Then
        frmColors.Show
        frmColors1.Checked = True
        Exit Sub
    End If
End Sub

Private Sub frmECopy_Click()
Clipboard.SetText (frmMain.txtNFODIZ.SelText)
End Sub

Private Sub frmEcut_Click()
Clipboard.SetText (frmMain.txtNFODIZ.SelText)
frmMain.txtNFODIZ.SelText = ""
End Sub

Private Sub frmEFind_Click()
frmFind.Show
End Sub

Private Sub frmEPaste_Click()
frmMain.txtNFODIZ.SelText = Clipboard.GetText
End Sub

Private Sub frmESelAll_Click()
frmMain.txtNFODIZ.SetFocus
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = Len(frmMain.txtNFODIZ.Text)
End Sub

Private Sub frmESelNone_Click()
frmMain.txtNFODIZ.SelStart = 0
frmMain.txtNFODIZ.SelLength = 0
End Sub

Private Sub frmExit_Click()
ans = MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "bizzyNFO4")
    If ans = vbYes Then
        ans2 = MsgBox("Do you want to save your " + frmInfo.lblFileType.Caption + " file before you exit?", vbYesNo + vbQuestion, "bizzyNFO4")
            If ans2 = vbYes Then
                If frmInfo.lblFileType.Caption = "NFO" Then
                    frmMDI.cd.DialogTitle = "Save an NFO file.."
                    frmMDI.cd.Filter = "bizzyNFO4 NFO File (*.nfo) | *.nfo"
                    frmMDI.cd.ShowSave
                    
                    frmMain.txtNFODIZ.SaveFile cd.FileName
                    MsgBox "Your file has been saved succesfully.", vbOKOnly + vbInformation, "bizzyNFO4"
                    End
                End If
                
                If frmInfo.lblFileType.Caption = "DIZ" Then
                    frmMDI.cd.DialogTitle = "Save an DIZ file.."
                    frmMDI.cd.Filter = "bizzyNFO4 DIZ File (*.diz) | *.diz"
                    frmMDI.cd.ShowSave
                
                    frmMain.txtNFODIZ.SaveFile ("" + cd.FileName + "")
                    MsgBox "Your file has been saved successfully.", vbOKOnly + vbInformation, "bizzyNFO4"
                    End
                End If
            Else
            End
            End If
    Else
    Cancel = 1
    End If
End Sub



Private Sub frmHAbout_Click()
frmAbout.Show
End Sub

Private Sub frmHHelp_Click()
frmHelp.Show
End Sub

Private Sub frmHTC_Click()
MsgBox "Help for Tool Chooser" & vbCrLf & vbCrLf & "1. Select a Tool from the Tools Window." & vbCrLf & "2. Select a Tool Number from the Set Tool Window." & vbCrLf & vbCrLf & "Use the new tool by pressing one of the following buttons." & vbCrLf & vbCrLf & "Tool 1: F1" & vbCrLf & "Tool 2: F2" & vbCrLf & "Tool 3: F3" & vbCrLf & "Tool 4: F4" & vbCrLf & "Tool 5: F5" & vbCrLf & "Tool 6: F7" & vbCrLf & "Tool 7: F8" & vbCrLf & vbCrLf & "Selected Tool from Tools Window: F6", vbOKOnly + vbInformation, "bizzyNFO4"

End Sub

Private Sub frmInfo1_Click()
    If frmInfo1.Checked = True Then
        frmInfo.Hide
        frmInfo1.Checked = False
        Exit Sub
    End If
    
    If frmInfo1.Checked = False Then
        frmInfo.Show
        frmInfo1.Checked = True
        Exit Sub
    End If
End Sub

Private Sub frmLDIZ_Click()
On Error GoTo 13
   
   If frmMain.txtNFODIZ.Text <> "" Then
        ans = MsgBox("Are you sure you want to load a new DIZ file? (This will erase all the contents in the bizzyNFO4 Workspace)", vbYesNo + vbQuestion, "bizzyNFO4")
            If ans = vbYes Then
                frmMDI.Caption = "bizzyNFO4 - Untitled.diz - 0 Character(s)"
                frmMain.txtNFODIZ.Text = ""
                cd.DialogTitle = "Open an DIZ file.."
                cd.Filter = "bizzyNFO4 DIZ File (*.diz) | *.diz"
                cd.ShowOpen
                frmMain.txtNFODIZ.LoadFile cd.FileName
                frmInfo.lblFileType.Caption = "DIZ"
                TitleOfND = cd.FileTitle
                LengthOfND = Len(frmMain.txtNFODIZ.Text)
                frmMDI.Caption = "bizzyNFO4 - " + TitleOfND + " - " + LengthOfND + " Character(s)"
            ElseIf ans = vbNo Then
                Exit Sub
            End If
    ElseIf frmMain.txtNFODIZ.Text = "" Then
        frmMDI.Caption = "bizzyNFO4 - Untitled.diz - 0 Character(s)"
        frmMain.txtNFODIZ.Text = ""
        cd.DialogTitle = "Open an DIZ file.."
        cd.Filter = "bizzyNFO4 DIZ File (*.diz) | *.diz"
        cd.ShowOpen
        frmMain.txtNFODIZ.LoadFile cd.FileName
        frmInfo.lblFileType.Caption = "DIZ"
        TitleOfND = cd.FileTitle
        LengthOfND = Len(frmMain.txtNFODIZ.Text)
        frmMDI.Caption = "bizzyNFO4 - " + TitleOfND + " - " + LengthOfND + " Character(s)"
    End If
13 End Sub

Private Sub frmLNFO_Click()
On Error GoTo 1

    If frmMain.txtNFODIZ.Text <> "" Then
        ans = MsgBox("Are you sure you want to load a new NFO file? (This will erase all the contents in the bizzyNFO4 Workspace)", vbYesNo + vbQuestion, "bizzyNFO4")
            If ans = vbYes Then
               frmMDI.Caption = "bizzyNFO4 - Untitled.nfo - 0 Character(s)"
                frmMain.txtNFODIZ.Text = ""
                cd.DialogTitle = "Open an NFO file.."
                cd.Filter = "bizzyNFO4 NFO File (*.nfo) | *.nfo"
                cd.ShowOpen
                frmMain.txtNFODIZ.LoadFile cd.FileName
                frmInfo.lblFileType.Caption = "NFO"
                TitleOfND = cd.FileTitle
                LengthOfND = Len(frmMain.txtNFODIZ.Text)
                frmMDI.Caption = "bizzyNFO4 - " + TitleOfND + " - " + LengthOfND + " Character(s)"
                Exit Sub
            ElseIf ans = vbNo Then
                Exit Sub
            End If
    ElseIf frmMain.txtNFODIZ.Text = "" Then
        cd.DialogTitle = "Open an NFO file.."
        cd.Filter = "bizzyNFO4 NFO File (*.nfo) | *.nfo"
        cd.ShowOpen
        frmMain.txtNFODIZ.LoadFile cd.FileName
        frmInfo.lblFileType.Caption = "NFO"
        TitleOfND = cd.FileTitle
        LengthOfND = Len(frmMain.txtNFODIZ.Text)
        frmMDI.Caption = "bizzyNFO4 - " + TitleOfND + " - " + LengthOfND + " Character(s)"
        Exit Sub
    End If
    
1 frmMDI.Caption = "bizzyNFO4 - Untitled.nfo - 0 Character(s)"

End Sub

Private Sub frmMain1_Click()
    If frmMain1.Checked = True Then
        frmMain.Hide
        frmMain1.Checked = False
        Exit Sub
    End If
    
    If frmMain1.Checked = False Then
        frmMain.Show
        frmMain1.Checked = True
        Exit Sub
    End If
End Sub

Private Sub frmNDIZ_Click()
Dim ans As Integer

ans = MsgBox("Are you sure you want to create a new DIZ file? (This will erase all the contents in the bizzyNFO4 Workspace)", vbYesNo + vbQuestion, "bizzyNFO4")
    If ans = vbYes Then
        frmMDI.Caption = "bizzyNFO4 - Untitled.diz - 0 Character(s)"
        frmMain.txtNFODIZ.Text = ""
        frmInfo.lblFileType.Caption = "DIZ"
    ElseIf ans = vbNo Then
        Exit Sub
    End If
End Sub

Private Sub frmNNFO_Click()
Dim ans As Integer

ans = MsgBox("Are you sure you want to create a new NFO file? (This will erase all the contents in the bizzyNFO4 Workspace)", vbYesNo + vbQuestion, "bizzyNFO4")
    If ans = vbYes Then
        frmMDI.Caption = "bizzyNFO4 - Untitled.nfo - 0 Character(s)"
        frmMain.txtNFODIZ.Text = ""
        frmInfo.lblFileType.Caption = "NFO"
    ElseIf ans = vbNo Then
        Exit Sub
    End If
End Sub

Private Sub frmOTDCLTO_Click()
    If frmOTDCLTO.Caption = "Turn Double Click ToolAdd On" Then '
        frmMain.TurnedOn.Caption = True
        frmOTDCLTO.Caption = "Turn Double Click ToolAdd Off"
        Exit Sub
    End If
    
    If frmOTDCLTO.Caption = "Turn Double Click ToolAdd Off" Then
        frmMain.TurnedOn.Caption = False
        frmOTDCLTO.Caption = "Turn Double Click ToolAdd On"
        Exit Sub
    End If
End Sub

Private Sub frmPNFODIZ_Click()
    If frmMain.txtNFODIZ.Text = "" Then
        MsgBox "You must have data to print if you want to print.", vbOKOnly + vbInformation, "bizzyNFO4"
    ElseIf frmMain.txtNFODIZ.Text <> "" Then
        Printer.FontName = frmMain.txtNFODIZ.Font.Name
        Printer.FontSize = frmMain.txtNFODIZ.Font.Size
        Printer.ForeColor = vbBlack
        Printer.Print frmMain.txtNFODIZ.Text
        Printer.EndDoc
    End If
        
End Sub

Private Sub frmSDIZ_Click()
On Error GoTo 23

cd.DialogTitle = "Save an DIZ File..."
cd.Filter = "bizzyNFO4 DIZ File (*.diz) | *.diz"
cd.ShowSave
TitleOfND = cd.FileTitle

    If cd.FileName <> "" Then
        frmMain.txtNFODIZ.SaveFile cd.FileName + ".diz"
        LengthOfND = Len(frmMain.txtNFODIZ.Text)
        TitleOfND = cd.FileTitle
        frmMDI.Caption = "bizzyNFO4 - " + TitleOfND + " - " + LengthOfND + " Character(s)"
      Exit Sub
    Else
    
    End If
23 End Sub

Private Sub frmSNFO_Click()
On Error GoTo 1

    cd.DialogTitle = "Save an NFO File..."
    cd.Filter = "bizzyNFO4 NFO File (*.nfo) | *.nfo"
    cd.ShowSave
    TitleOfND = cd.FileTitle
    
    If cd.FileName <> "" Then
    frmMain.txtNFODIZ.SaveFile cd.FileName + ".nfo"
    LengthOfND = Len(frmMain.txtNFODIZ.Text)
    TitleOfND = cd.FileTitle
    frmMDI.Caption = "bizzyNFO4 - " + TitleOfND + " - " + LengthOfND + " Character(s)"
   Exit Sub
    Else
    
    End If



1 End Sub

Private Sub frmTChooser1_Click()
    If frmTChooser1.Checked = True Then
        frmTChooser.Hide
        frmTChooser1.Checked = False
        Exit Sub
    End If
    
    If frmTChooser1.Checked = False Then
        frmTChooser.Show
        frmTChooser1.Checked = True
        Exit Sub
    End If
End Sub

Private Sub frmTools1_Click()
     If frmTools1.Checked = True Then
        frmTools.Hide
        frmTools1.Checked = False
        Exit Sub
    End If
    
    If frmTools1.Checked = False Then
        frmTools.Show
        frmTools1.Checked = True
        Exit Sub
    End If
    
End Sub



Private Sub MDIForm_Load()


frmColors.Show
frmColors.Top = 195
frmColors.Left = 20

frmInfo.Show
frmInfo.Top = 8000
frmInfo.Left = 20

frmTools.Show
frmTools.Top = 195
frmTools.Left = frmMDI.Width - 3090
frmTools.optChar(0).Value = True

frmInfo.Top = 7000

frmTChooser.Show
frmTChooser.Left = frmMDI.Width - 3090
frmTChooser.Top = 9250
frmTChooser.optSel(0).Value = True

frmMain.Show
frmMain.Top = 195
frmMain.Left = 2865
frmMain.Height = frmColors.Height + frmInfo.Height + 40

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim ans As Integer
Dim ans2 As Integer

ans = MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "bizzyNFO4")

    If ans = vbYes Then
        ans2 = MsgBox("Do you want to save your " + frmInfo.lblFileType.Caption + " file before you exit?", vbYesNo + vbQuestion, "bizzyNFO4")
            If ans2 = vbYes Then
                If frmInfo.lblFileType.Caption = "NFO" Then
                    frmMDI.cd.DialogTitle = "Save an NFO file.."
                    frmMDI.cd.Filter = "bizzyNFO4 NFO File (*.nfo) | *.nfo"
                    frmMDI.cd.ShowSave
                    
                    frmMain.txtNFODIZ.SaveFile cd.FileName
                    MsgBox "Your file has been saved succesfully.", vbOKOnly + vbInformation, "bizzyNFO4"
                    End
                End If
                
                If frmInfo.lblFileType.Caption = "DIZ" Then
                    frmMDI.cd.DialogTitle = "Save an DIZ file.."
                    frmMDI.cd.Filter = "bizzyNFO4 DIZ File (*.diz) | *.diz"
                    frmMDI.cd.ShowSave
                
                    frmMain.txtNFODIZ.SaveFile ("" + cd.FileName + "")
                    MsgBox "Your file has been saved successfully.", vbOKOnly + vbInformation, "bizzyNFO4"
                    End
                End If
            Else
            End
            End If
    Else
    Cancel = 1
    End If

End Sub

