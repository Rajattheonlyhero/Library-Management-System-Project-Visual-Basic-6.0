VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Welcome"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   30
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   12480
      Top             =   5880
   End
   Begin VB.Label LblLib2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Books,,,, Your Best Companion"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2925
      Left            =   3840
      TabIndex        =   5
      Top             =   5760
      Width           =   7260
   End
   Begin VB.Label LblRpbc2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Project"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2820
      Left            =   2280
      TabIndex        =   4
      Top             =   2880
      Width           =   6240
   End
   Begin VB.Label LblWelcome2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To The"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4020
      Left            =   4560
      TabIndex        =   3
      Top             =   -120
      Width           =   2220
   End
   Begin VB.Label LblLib1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Meet Your Next Favorite Books"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2685
      Left            =   2160
      TabIndex        =   2
      Top             =   4800
      Width           =   6840
   End
   Begin VB.Label LblRpbc1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRARY MANAGEMENT"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   8265
      Left            =   -2175
      TabIndex        =   1
      Top             =   -600
      Width           =   16380
   End
   Begin VB.Label LblWelcome1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   5355
      Left            =   3720
      TabIndex        =   0
      Top             =   -1920
      Width           =   3870
   End
   Begin VB.Image Img1 
      Height          =   8520
      Left            =   120
      Picture         =   "frmWelcome.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11640
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

    Img1.Height = Me.ScaleHeight
    Img1.Width = Me.ScaleWidth
    LblWelcome1.Left = (Me.ScaleWidth / 2 - LblWelcome1.Width / 2) - 120
    LblWelcome2.Left = Me.ScaleWidth / 2 - LblWelcome2.Width / 2

    LblRpbc1.Left = (Me.ScaleWidth / 2 - LblRpbc1.Width / 2) + 60
    LblRpbc2.Left = (Me.ScaleWidth / 2 - LblRpbc1.Width / 2)
    
    LblLib1.Left = Me.ScaleWidth / 2 - LblLib1.Width / 2
    LblLib2.Left = (Me.ScaleWidth / 2 - LblLib2.Width / 2) + 180
    
End Sub


Private Sub LblRpbc1_Click()
    Frmlogin.Show
    Unload Me
    
    Timer1.Enabled = False
End Sub

Private Sub LblWelcome1_Click()
   Frmlogin.Show
   Unload Me
   
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Frmlogin.Show
    Unload Me
    Timer1.Enabled = False
End Sub


