VERSION 5.00
Begin VB.Form frmTools 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tools"
   ClientHeight    =   8625
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
   ScaleHeight     =   8625
   ScaleWidth      =   2565
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ô"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   84
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   5160
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ó"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   83
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   5160
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ò"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   82
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   5160
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ñ"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   81
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   5160
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ð"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   80
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   5160
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ï"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   79
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   4920
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "î"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   78
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   4920
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "í"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   77
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   4920
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ì"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   76
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   4920
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ë"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   75
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   4920
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   74
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   4680
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   73
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   4680
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   72
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   4680
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ç"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   71
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   4680
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "æ"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   70
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   4680
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "å"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   69
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   4440
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ä"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   68
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   4440
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ã"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   67
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   4440
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "â"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   66
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   4440
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "á"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   65
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   4440
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "à"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   64
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   4200
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ß"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   63
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   4200
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Þ"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   62
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   4200
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ý"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   61
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   4200
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ü"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   60
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   4200
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Û"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   59
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   3960
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ú"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   58
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   3960
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ù"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   57
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   3960
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ø"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   56
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   3960
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   55
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   3960
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ö"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   54
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3720
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Õ"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   53
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3720
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ô"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   52
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3720
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ó"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   51
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   3720
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ò"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   50
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   3720
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ñ"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   49
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3480
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ð"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   48
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3480
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ï"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   47
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3480
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Î"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   46
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   3480
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Í"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   45
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   3480
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ì"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   44
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   3240
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ë"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   43
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   3240
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ê"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   42
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3240
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "É"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   41
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   3240
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "È"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   40
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   3240
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ç"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   39
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Æ"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   38
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Å"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ä"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   36
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ã"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   35
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Â"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   34
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Á"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "À"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¿"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¾"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "½"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2520
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¼"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2520
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "»"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2520
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "º"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2520
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¹"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2520
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¸"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2280
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "·"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2280
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¶"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2280
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "µ"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2280
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "´"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2280
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "³"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2040
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   94
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6480
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   93
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6480
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   92
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6480
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   91
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6480
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   90
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6480
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   89
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6240
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   88
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6240
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   87
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6240
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "²"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2040
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "±"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2040
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2040
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¯"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2040
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "®"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¬"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "«"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ª"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "©"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1560
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¨"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1560
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "§"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¦"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1560
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¥"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¤"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¢"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¡"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¡"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   86
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   255
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   85
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label lblCustom1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Custom Characters"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label lblBigPreview 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "HyperFont"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label lblPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Big Preview"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Shape shpBackGround3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1095
      Left            =   -120
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label lblChars 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Characters"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblTools 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tools"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1455
      Left            =   -120
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Shape shpBackGround2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   4695
      Left            =   -120
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "frmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KEYZ As String

Private Sub optChar_Click(Index As Integer)
    For i = 0 To 84
        If optChar(i).Value = True Then
            lblBigPreview.Caption = "" + optChar(i).Caption + ""
            frmInfo.lblSelectedTool.Caption = "" + optChar(i).Caption + ""
            KEYZ = optChar(i).Index + 160
            frmInfo.lblKeyStroke.Caption = "ALT+0" + KEYZ + ""
            Exit For
        End If
    Next i
    
    For i2 = 85 To 94
        If optChar(i2).Value = True Then
            lblBigPreview.Caption = "" + optChar(i2).Caption + ""
            frmInfo.lblSelectedTool.Caption = "" + optChar(i2).Caption + ""
            frmInfo.lblKeyStroke.Caption = "Custom"
            Exit For
        End If
    Next i2
End Sub

