VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmBookSu 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Book Submission"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14955
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   14955
   WindowState     =   2  'Maximized
   Begin VB.Frame Book 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Book Action Area"
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
      Height          =   10215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14295
      Begin VB.CommandButton Command2 
         Caption         =   "Confirm Submit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6480
         TabIndex        =   28
         Top             =   7200
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Back To Main Screen"
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
         Left            =   240
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   7320
         Width           =   2535
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Submit New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4680
         TabIndex        =   25
         Top             =   7200
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Search Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8655
         Left            =   8280
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   5775
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   495
            Left            =   120
            Top             =   3000
            Visible         =   0   'False
            Width           =   5415
            _ExtentX        =   9551
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
            Connect         =   $"FrmBookSu.frx":0000
            OLEDBString     =   $"FrmBookSu.frx":0095
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "select * from submit_details"
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "FrmBookSu.frx":012A
            Height          =   2055
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Visible         =   0   'False
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   3625
            _Version        =   393216
            BackColor       =   12648384
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
         Begin VB.Label Label15 
            Caption         =   "Records :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   2040
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Submission Details:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   3735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Member Details"
         Height          =   2415
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   5655
         Begin VB.TextBox TextCodeMem 
            Height          =   495
            Left            =   2640
            TabIndex        =   18
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox TextMemName 
            Height          =   495
            Left            =   2640
            TabIndex        =   17
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Member's Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   1320
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Book Details"
         Height          =   3615
         Left            =   240
         TabIndex        =   1
         Top             =   3360
         Width           =   7695
         Begin VB.TextBox TextCodeBook 
            Height          =   495
            Left            =   3120
            TabIndex        =   9
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox TextBookName 
            Height          =   495
            Left            =   3120
            TabIndex        =   8
            Top             =   1440
            Width           =   2415
         End
         Begin VB.ComboBox CmbDay1 
            Height          =   315
            Left            =   3120
            TabIndex        =   7
            Text            =   "Select"
            Top             =   2400
            Width           =   855
         End
         Begin VB.ComboBox CmbMonth1 
            Height          =   315
            Left            =   3960
            TabIndex        =   6
            Text            =   "Select"
            Top             =   2400
            Width           =   855
         End
         Begin VB.ComboBox CmbYear1 
            Height          =   315
            Left            =   4800
            TabIndex        =   5
            Text            =   "Select"
            Top             =   2400
            Width           =   855
         End
         Begin VB.ComboBox CmbDay2 
            Height          =   315
            Left            =   3120
            TabIndex        =   4
            Text            =   "Select"
            Top             =   3000
            Width           =   855
         End
         Begin VB.ComboBox CmbMonth2 
            Height          =   315
            Left            =   3960
            TabIndex        =   3
            Text            =   "Select"
            Top             =   3000
            Width           =   855
         End
         Begin VB.ComboBox CmbYear2 
            Height          =   315
            Left            =   4800
            TabIndex        =   2
            Text            =   "Select"
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Book Title"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Issue Date:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0FFC0&
            Caption         =   "DD-MM-YYYY"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   12
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Submission Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   3000
            Width           =   2895
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "DD-MM-YYYY"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   10
            Top             =   3000
            Width           =   1815
         End
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   735
         Left            =   120
         TabIndex        =   26
         Top             =   9360
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   1296
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   5
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   14261
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
End
Attribute VB_Name = "FrmBookSu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdadd_Click()
Adodc1.Visible = False
DataGrid1.Visible = True
Adodc1.Recordset.AddNew
End Sub

Private Sub Command1_Click()
MDIfrm.Show
Unload Me

End Sub

Private Sub Command2_Click()
Dim dt1, dt2 As String
If TextCodeMem.Text = "" Or TextCodeBook.Text = "" Or TextMemName.Text = "" Or TextBookName.Text = "" Then
MsgBox "Enter all compulsory information.", vbInformation, "Book Entry"
            Exit Sub
    End If
dt1 = CmbDay1.Text & "/" & CmbMonth1.Text & "/" & CmbYear1.Text
dt2 = CmbDay2.Text & "/" & CmbMonth2.Text & "/" & CmbYear2.Text
Adodc1.Recordset.Fields("Member_Code").Value = TextCodeMem.Text
Adodc1.Recordset.Fields("Member_Name").Value = TextMemName.Text
Adodc1.Recordset.Fields("Book_Code").Value = TextCodeBook.Text
Adodc1.Recordset.Fields("Book_Name").Value = TextBookName.Text
Adodc1.Recordset.Fields("Isu_Dt").Value = dt1
Adodc1.Recordset.Fields("Sub_Dt").Value = dt2
Adodc1.Recordset.Update
MsgBox "Record succesfully updated", vbOKOnly, App.Title

End Sub

Private Sub Form_Load()
StatusBar1.Panels(1) = "Current User : " & Book1.userNm & ""
End Sub
