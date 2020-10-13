VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUserMng 
   BackColor       =   &H00C0FFFF&
   Caption         =   "User Management"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   13455
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "User Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6135
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   8415
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00FFFF00&
         Caption         =   "Go Back"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5280
         Width           =   2415
      End
      Begin VB.CommandButton CmdDeleteAcc 
         BackColor       =   &H0080C0FF&
         Caption         =   "Delete Account"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3840
         Width           =   5895
      End
      Begin VB.CommandButton CmdEditAcc 
         BackColor       =   &H0080C0FF&
         Caption         =   "Edit Your Account"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2400
         Width           =   5895
      End
      Begin VB.CommandButton CmdCreateAcc 
         BackColor       =   &H0080C0FF&
         Caption         =   "Create a New Account"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   5895
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   6960
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11933
            Object.ToolTipText     =   "Current User"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "5/20/2019"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "2:31 PM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmUserMng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
Frmlogin.Show
Unload Me
End Sub

Private Sub CmdCreateAcc_Click()
FrmCreateAcc.Show
Unload Me
End Sub

Private Sub CmdDeleteAcc_Click(Index As Integer)
FrmAdminLogin.Show
Unload Me
End Sub

Private Sub CmdEditAcc_Click(Index As Integer)
Frmeditacc.Show
Unload Me

End Sub

Private Sub Form_Load()
StatusBar1.Panels(1) = "Current User : " & Book1.userNm & ""

End Sub
