VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "bizzyNFO4 - Splash"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   2550
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   960
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub Form_Click()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
a = Val(a) + 1
frmSplash.Top = frmSplash.Top - 100
    
    If a = 120 Then
    frmMainMenu.Show
    Unload Me
    End If
End Sub
