VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmEditAcc 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Change User Name or Password"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   13470
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Change Details "
      Height          =   8655
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   9735
      Begin VB.CommandButton Command2 
         Caption         =   "Update  Your Credentials"
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
         Left            =   4920
         TabIndex        =   13
         Top             =   6360
         Width           =   4695
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   6360
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   5400
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   4320
         Width           =   4575
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmEditAcc.frx":0000
         Height          =   1215
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Visible         =   0   'False
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2143
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Admin Credentials"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TextSearch 
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find Your Credentials"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   4695
      End
      Begin VB.CommandButton Command6 
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7440
         Width           =   4695
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   720
         Top             =   8040
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   873
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
         Connect         =   $"FrmEditAcc.frx":0015
         OLEDBString     =   $"FrmEditAcc.frx":00AA
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Your New Password :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   6000
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Type Your New Password :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   5040
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Type Your New User Name :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3960
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter Your Username To check For Your Credentials :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   7575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   735
      Left            =   1080
      TabIndex        =   14
      Top             =   9960
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6853
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
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Edit Existing User Account :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   9735
   End
End
Attribute VB_Name = "Frmeditacc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If TextSearch.Text = "" Then
MsgBox "Please provide input", vbOKOnly, App.Title
Exit Sub
End If
DataGrid1.Visible = True
Adodc1.RecordSource = "select * from userbase where Username='" & TextSearch.Text & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Record not found", vbOKOnly, App.Title
Exit Sub
Else
Adodc1.Caption = Adodc1.RecordSource
Label1.Visible = True
Label3.Visible = True
Label5.Visible = True
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Command2.Visible = True
Text1.Text = Adodc1.Recordset.Fields("Username").Value
Text2.Text = Adodc1.Recordset.Fields("Password").Value
End If
End Sub

Private Sub Command2_Click()
DataGrid1.Visible = False
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("Username").Value = Text1.Text
If Text2.Text = Text3.Text Then
Adodc1.Recordset.Fields("Password").Value = Text2.Text
Else
MsgBox "User Password And Confirm Password Not same,, Plzz Try Again", vbCritical, App.Title
End If
Adodc1.Recordset.Update
MsgBox "Record succesfully updated", vbOKOnly, App.Title
Exit Sub
End Sub

Private Sub Command6_Click()
FrmUserMng.Show
Unload Me
End Sub

Private Sub Form_Load()
StatusBar1.Panels(1) = "Current User : " & Book1.userNm & ""
End Sub
