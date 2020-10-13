VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmLogin 
   BackColor       =   &H000080FF&
   Caption         =   "Login"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12420
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc loginado 
      Height          =   735
      Left            =   2160
      Top             =   9360
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmlogin.frx":0000
      OLEDBString     =   $"frmlogin.frx":0095
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from userbase"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "User Login"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7215
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   9015
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF80&
         Caption         =   "Edit Existing"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6240
         Width           =   2535
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Exit"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4440
         TabIndex        =   10
         Top             =   3720
         Width           =   2625
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Admin Sign Up"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5160
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5160
         Width           =   2535
      End
      Begin VB.CommandButton CmdIsuDtl 
         BackColor       =   &H00FFFF80&
         Caption         =   "Book Issue Rules"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox txtpass 
         Height          =   555
         IMEMode         =   3  'DISABLE
         Left            =   4080
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2760
         Width           =   3285
      End
      Begin VB.CommandButton cmdlogin 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         TabIndex        =   2
         Top             =   3720
         Width           =   2625
      End
      Begin VB.TextBox txtuser 
         Height          =   585
         Left            =   4080
         TabIndex        =   1
         Top             =   1800
         Width           =   3285
      End
      Begin VB.Label Username 
         BackStyle       =   0  'Transparent
         Caption         =   "Wanna Change Credentials?? "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   6600
         Width           =   4200
      End
      Begin VB.Label Username 
         BackStyle       =   0  'Transparent
         Caption         =   "Already Have A Account??"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   6240
         Width           =   3480
      End
      Begin VB.Label Username 
         BackStyle       =   0  'Transparent
         Caption         =   "New User, Sign Up Here:"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   5160
         TabIndex        =   8
         Top             =   4680
         Width           =   3480
      End
      Begin VB.Label Username 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Library Rules Here:"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   4680
         Width           =   3480
      End
      Begin VB.Label Password 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   2760
         Width           =   2640
      End
      Begin VB.Label Username 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   1800
         Width           =   2640
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   735
      Left            =   1560
      TabIndex        =   14
      Top             =   8280
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5583
            Object.ToolTipText     =   "Current User"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "5/20/2019"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:48 PM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock On"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Num Lock On"
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
Attribute VB_Name = "Frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
Dim x As String
x = MsgBox("Do you really want to exit", vbYesNo + vbCritical, "Delete Confirmation")
If x = vbYes Then
End
End If
End Sub

Private Sub CmdIsuDtl_Click()
FrmBookIsuRu.Show
Unload Me

End Sub

Private Sub cmdlogin_Click()

If txtuser.Text = "" Or txtpass.Text = "" Then
MsgBox "Plzz Fill In the Details", vbCritical, App.Title
    Exit Sub
    End If

loginado.RecordSource = "select * from userbase where Username ='" + txtuser.Text + "' and Password = '" + txtpass.Text + "'"


loginado.Refresh

If loginado.Recordset.EOF Then
MsgBox "Login failed,,,Please Enter Correct Usename And Password!!!", vbCritical, App.Title
Else

MsgBox "Login Succesful", vbInformation, App.Title
Book1.userNm = txtuser.Text
MDIfrm.Show
End If

End Sub

Private Sub Command1_Click()
FrmCreateAcc.Show
Unload Me
End Sub

Private Sub Command2_Click()
FrmUserMng.Show
Unload Me
End Sub

Private Sub Form_Load()
StatusBar1.Panels(1) = "Current User : " & Book1.userNm & ""
End Sub
