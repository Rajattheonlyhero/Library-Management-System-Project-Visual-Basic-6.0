VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmMember 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Member Operations"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   14190
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9600
      Top             =   9240
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   582
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
      Connect         =   $"FrmMember.frx":0000
      OLEDBString     =   $"FrmMember.frx":0095
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from member_details"
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
   Begin VB.Frame FrmMember 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Member Action Area"
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
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   15375
      Begin VB.TextBox TextFine 
         Height          =   495
         Left            =   1800
         TabIndex        =   38
         Top             =   7680
         Width           =   2415
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
         Height          =   975
         Left            =   7680
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   8280
         Width           =   1335
      End
      Begin VB.Frame Frame1 
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
         Left            =   9480
         TabIndex        =   30
         Top             =   360
         Width           =   5775
         Begin VB.CommandButton Command2 
            Caption         =   "Search"
            Height          =   855
            Left            =   120
            TabIndex        =   39
            Top             =   2040
            Width           =   1695
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "FrmMember.frx":012A
            Height          =   2055
            Left            =   120
            TabIndex        =   34
            Top             =   3960
            Visible         =   0   'False
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   3625
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
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   3480
            TabIndex        =   33
            Text            =   "Select"
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox TextSearch 
            Height          =   495
            Left            =   3240
            TabIndex        =   31
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label13 
            Caption         =   "Select Search Category:"
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
            TabIndex        =   40
            Top             =   480
            Width           =   3015
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
            TabIndex        =   35
            Top             =   3480
            Width           =   3015
         End
         Begin VB.Label Label14 
            Caption         =   "Searching Word:"
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
            TabIndex        =   32
            Top             =   1200
            Width           =   3015
         End
      End
      Begin VB.OptionButton OptFemale 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Female"
         Height          =   495
         Left            =   6000
         TabIndex        =   29
         Top             =   7200
         Width           =   1575
      End
      Begin VB.OptionButton OptMale 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Male"
         Height          =   615
         Left            =   6000
         TabIndex        =   28
         Top             =   6600
         Width           =   1575
      End
      Begin VB.TextBox TextContact 
         Height          =   495
         Left            =   1800
         TabIndex        =   27
         Top             =   6840
         Width           =   2415
      End
      Begin VB.TextBox Textfee 
         Height          =   495
         Left            =   6480
         TabIndex        =   24
         Top             =   5880
         Width           =   2415
      End
      Begin VB.TextBox TextCity 
         Height          =   495
         Left            =   1800
         TabIndex        =   21
         Top             =   5880
         Width           =   2415
      End
      Begin VB.TextBox TextAddress 
         Height          =   1695
         Left            =   1800
         TabIndex        =   20
         Top             =   3720
         Width           =   4455
      End
      Begin VB.ComboBox CmbYear 
         Height          =   315
         Left            =   3480
         TabIndex        =   17
         Text            =   "Select"
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox CmbMonth 
         Height          =   315
         Left            =   2640
         TabIndex        =   16
         Text            =   "Select"
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox CmbDay 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Text            =   "Select"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox TextFather 
         Height          =   495
         Left            =   6600
         TabIndex        =   14
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox TextLast 
         Height          =   495
         Left            =   3960
         TabIndex        =   13
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox TextFirst 
         Height          =   495
         Left            =   1320
         TabIndex        =   12
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox TextCode 
         Height          =   495
         Left            =   1320
         TabIndex        =   11
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "Delete"
         Height          =   855
         Left            =   5880
         TabIndex        =   4
         Top             =   8400
         Width           =   1695
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "Update"
         Height          =   855
         Left            =   4080
         TabIndex        =   3
         Top             =   8400
         Width           =   1695
      End
      Begin VB.CommandButton cmddisplay 
         Caption         =   "Display"
         Height          =   855
         Left            =   2160
         TabIndex        =   2
         Top             =   8400
         Width           =   1695
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add New"
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   8400
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fine:"
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
         TabIndex        =   37
         Top             =   7680
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Gender"
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
         Left            =   4560
         TabIndex        =   26
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Contact No."
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
         TabIndex        =   25
         Top             =   6840
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Membership Fee"
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
         Left            =   4320
         TabIndex        =   23
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "City"
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
         TabIndex        =   22
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Address"
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
         Top             =   3840
         Width           =   1215
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
         Left            =   4440
         TabIndex        =   18
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Father's Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   10
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Join Date"
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
         TabIndex        =   7
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Name"
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
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   735
      Left            =   120
      TabIndex        =   41
      Top             =   9960
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16801
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
Attribute VB_Name = "FrmMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dt As String, sex As String
Private Sub cmdadd_Click()
Adodc1.Visible = False
DataGrid1.Visible = True
Adodc1.Recordset.AddNew
End Sub

Private Sub cmddelete_Click()
Adodc1.Recordset.Delete
MsgBox ("Record successfully deleted"), vbOKOnly, App.Title
Exit Sub
End Sub

Private Sub cmddisplay_Click()
DataGrid1.Visible = True
Adodc1.RecordSource = "select * from member_details"
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub cmdupdate_Click()
If TextCode = "" Or TextFirst = "" Or TextFather = "" Or TextLast = "" Or _
        TextAddress = "" Or TextCity = "" Or Textfee = "" Then
            MsgBox "Enter all compulsory information.", vbInformation, "Member Entry"
            Exit Sub
    End If
    dt = CmbDay.Text & "/" & CmbMonth.Text & "/" & CmbYear.Text
    If OptMale.Value = True Then
        sex = "M"
    Else
        sex = "F"
    End If


Adodc1.Recordset.Fields("Code").Value = Val(TextCode.Text)
Adodc1.Recordset.Fields("Surname").Value = TextLast.Text
Adodc1.Recordset.Fields("First_Name").Value = TextFirst.Text
Adodc1.Recordset.Fields("Father Name").Value = TextFather.Text
Adodc1.Recordset.Fields("Join_Dt").Value = dt
Adodc1.Recordset.Fields("Address").Value = TextAddress.Text
Adodc1.Recordset.Fields("City").Value = TextCity.Text
Adodc1.Recordset.Fields("Cnt_No").Value = TextContact.Text
Adodc1.Recordset.Fields("Fee").Value = Textfee.Text
Adodc1.Recordset.Fields("Gender").Value = sex
Adodc1.Recordset.Fields("Fine").Value = Val(TextFine.Text)

Adodc1.Recordset.Update
MsgBox "Record succesfully updated", vbOKOnly, App.Title
TextCode.Text = ""
TextFirst.Text = ""
TextLast.Text = ""
TextFather.Text = ""
TextAddress.Text = ""
TextCity.Text = ""
TextContact.Text = ""
Textfee.Text = ""
TextFine.Text = ""
CmbDay.Text = ""
CmbMonth.Text = ""
CmbYear.Text = ""
Exit Sub
End Sub

Private Sub Command1_Click()
MDIfrm.Show
Unload Me

End Sub

Private Sub Command2_Click()
If Combo4.Text = "" Then
MsgBox "Please Select Category For Searching", vbOKOnly, App.Title
End If
If Combo4.Text = "First Name" Then
DataGrid1.Visible = True
If TextSearch.Text = "" Then
MsgBox "Please provide input", vbOKOnly, App.Title
Exit Sub
End If
Adodc1.RecordSource = "select * from member_details where First_Name='" & TextSearch.Text & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Record not found", vbOKOnly, App.Title
Exit Sub
Else
Adodc1.Caption = Adodc1.RecordSource
End If
Else
DataGrid1.Visible = True
If TextSearch.Text = "" Then
MsgBox "Please provide input", vbOKOnly, App.Title
Exit Sub
End If
Adodc1.RecordSource = "select * from member_details where Code='" & TextSearch.Text & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Record not found", vbOKOnly, App.Title
Exit Sub
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End If
End Sub

Private Sub Form_Load()
StatusBar1.Panels(1) = "Current User : " & Book1.userNm & ""
    Dim i As Integer
    
    'DAY COMBO
    For i = 1 To 31
        CmbDay.AddItem i
    Next
    'MONTH COMBO
    For i = 1 To 12
        CmbMonth.AddItem i
    Next
    'YEAR COMBO
    For i = 1950 To 2050
        CmbYear.AddItem i
    Next
    Combo4.AddItem "Code"
    Combo4.AddItem "First Name"
End Sub

