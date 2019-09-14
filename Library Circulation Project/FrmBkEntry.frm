VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmBkEntry 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Book Operations"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14865
   FillColor       =   &H00C0FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   14865
   WindowState     =   2  'Maximized
   Begin VB.Frame Book 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Book Entry"
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
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   14175
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
         Left            =   360
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   8280
         Width           =   2535
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add New"
         Height          =   855
         Left            =   360
         TabIndex        =   20
         Top             =   7080
         Width           =   1695
      End
      Begin VB.CommandButton cmddisplay 
         Caption         =   "Display"
         Height          =   855
         Left            =   2160
         TabIndex        =   19
         Top             =   7080
         Width           =   1695
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "Update"
         Height          =   855
         Left            =   3960
         TabIndex        =   18
         Top             =   7080
         Width           =   1695
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "Delete"
         Height          =   855
         Left            =   5760
         TabIndex        =   17
         Top             =   7080
         Width           =   1695
      End
      Begin VB.TextBox TextCode 
         Height          =   495
         Left            =   3120
         TabIndex        =   16
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox TextTitle 
         Height          =   495
         Left            =   3120
         TabIndex        =   15
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox TextAuthor 
         Height          =   495
         Left            =   3120
         TabIndex        =   14
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox TextPublisher 
         Height          =   495
         Left            =   3120
         TabIndex        =   13
         Top             =   3840
         Width           =   2415
      End
      Begin VB.ComboBox CmbDay 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Text            =   "Select"
         Top             =   4440
         Width           =   855
      End
      Begin VB.ComboBox CmbMonth 
         Height          =   315
         Left            =   2760
         TabIndex        =   11
         Text            =   "Select"
         Top             =   4440
         Width           =   855
      End
      Begin VB.ComboBox CmbYear 
         Height          =   315
         Left            =   3600
         TabIndex        =   10
         Text            =   "Select"
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox TextPrice 
         Height          =   495
         Left            =   3000
         TabIndex        =   9
         Top             =   5160
         Width           =   2775
      End
      Begin VB.TextBox TextFrom 
         Height          =   495
         Left            =   3000
         TabIndex        =   8
         Top             =   5880
         Width           =   2775
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
         Left            =   7920
         TabIndex        =   1
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton Command2 
            Caption         =   "Search"
            Height          =   855
            Left            =   120
            TabIndex        =   31
            Top             =   1920
            Width           =   1695
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   120
            Top             =   5400
            Visible         =   0   'False
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   661
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
            Connect         =   $"FrmBkEntry.frx":0000
            OLEDBString     =   $"FrmBkEntry.frx":0095
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "select * from book_Database"
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
         Begin VB.TextBox TextSearch 
            Height          =   495
            Left            =   3240
            TabIndex        =   4
            Top             =   1200
            Width           =   2415
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   3480
            TabIndex        =   3
            Text            =   "Select"
            Top             =   480
            Width           =   2055
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "FrmBkEntry.frx":012A
            Height          =   2055
            Left            =   120
            TabIndex        =   2
            Top             =   3240
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
            TabIndex        =   7
            Top             =   480
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
            TabIndex        =   6
            Top             =   1200
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
            TabIndex        =   5
            Top             =   2880
            Width           =   3015
         End
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   9240
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   1296
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   5
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   14049
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
         TabIndex        =   29
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Title"
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
         TabIndex        =   28
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Purch. Date"
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
         TabIndex        =   27
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Author:"
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
         Left            =   240
         TabIndex        =   26
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Publisher:"
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
         Left            =   240
         TabIndex        =   25
         Top             =   3840
         Width           =   2415
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
         TabIndex        =   24
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Price :"
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
         TabIndex        =   23
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "From Optional :"
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
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FrmBkEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dt As String
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
Adodc1.RecordSource = "select * from book_database"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub
Private Sub cmdupdate_Click()
If TextCode = "" Or TextFrom = "" Or TextTitle = "" Or TextAuthor = "" Or TextPublisher.Text = "" Or TextPrice.Text = "" Then
            MsgBox "Enter all compulsory information.", vbInformation, "Book Entry"
            Exit Sub
    End If
    dt = CmbDay.Text & "/" & CmbMonth.Text & "/" & CmbYear.Text


Adodc1.Recordset.Fields("Code").Value = Val(TextCode.Text)
Adodc1.Recordset.Fields("Title").Value = TextTitle.Text
Adodc1.Recordset.Fields("Author").Value = TextAuthor.Text
Adodc1.Recordset.Fields("Publisher").Value = TextPublisher.Text
Adodc1.Recordset.Fields("Pur_Dt").Value = dt
Adodc1.Recordset.Fields("Pur_From").Value = TextFrom.Text
Adodc1.Recordset.Fields("price").Value = TextPrice.Text
Adodc1.Recordset.Update
MsgBox "Record succesfully updated", vbOKOnly, App.Title
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
If Combo4.Text = "Title" Then
DataGrid1.Visible = True
If TextSearch.Text = "" Then
MsgBox "Please provide input", vbOKOnly, App.Title
Exit Sub
End If
Adodc1.RecordSource = "select * from book_database where Title='" & TextSearch.Text & "'"
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
Adodc1.RecordSource = "select * from book_database where Author='" & TextSearch.Text & "'"
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
    
    Combo4.AddItem "Title"
    Combo4.AddItem "Author"
    
End Sub
