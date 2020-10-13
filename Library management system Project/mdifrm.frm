VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form MDIfrm 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Main Screen"
   ClientHeight    =   8370
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   14325
   WindowState     =   2  'Maximized
   Begin VB.Frame menuframe 
      BackColor       =   &H0080FFFF&
      Caption         =   "Decide Your Journey"
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
      Height          =   9615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   8760
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   1296
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   5
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   14896
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
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00FFFF80&
         Caption         =   "Exit"
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
         Left            =   5520
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7560
         Width           =   2895
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Journey to The Oceans of Books"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   9600
         TabIndex        =   5
         Top             =   3360
         Width           =   3975
         Begin VB.CommandButton CmdBookisland 
            BackColor       =   &H00FFFF80&
            Caption         =   "Book Island"
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
            Left            =   600
            MaskColor       =   &H00C0FFC0&
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2280
            Width           =   2895
         End
         Begin VB.CommandButton CmdBkRpt 
            BackColor       =   &H00FFFF80&
            Caption         =   "Books In Trend"
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
            Left            =   600
            MaskColor       =   &H00C0FFC0&
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1320
            Width           =   2895
         End
         Begin VB.CommandButton CmdBkEntry 
            BackColor       =   &H00FFFF80&
            Caption         =   "Books Action Area"
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
            Left            =   600
            MaskColor       =   &H00C0FFC0&
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Issue/Submission"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   4920
         TabIndex        =   4
         Top             =   3000
         Width           =   3975
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFF80&
            Caption         =   "Submit"
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
            Left            =   2040
            MaskColor       =   &H00C0FFC0&
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmduserbase 
            BackColor       =   &H00FFFF80&
            Caption         =   "UserBase"
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
            Left            =   600
            MaskColor       =   &H00C0FFC0&
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3240
            Width           =   2895
         End
         Begin VB.CommandButton CmdIsuRpt 
            BackColor       =   &H00FFFF80&
            Caption         =   "Issue/Submit Report"
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
            Left            =   600
            MaskColor       =   &H00C0FFC0&
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2400
            Width           =   2895
         End
         Begin VB.CommandButton CmdBkSubISu 
            BackColor       =   &H00FFFF80&
            Caption         =   "Issue"
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
            Left            =   600
            MaskColor       =   &H00C0FFC0&
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1440
            Width           =   1455
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
            Left            =   600
            MaskColor       =   &H00C0FFC0&
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   480
            Width           =   2895
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0080FFFF&
         Caption         =   "The Member's Corner"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   360
         TabIndex        =   3
         Top             =   3600
         Width           =   3975
         Begin VB.CommandButton CmdMbrRpt 
            BackColor       =   &H00FFFF80&
            Caption         =   "Member Island"
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
            Left            =   480
            MaskColor       =   &H00C0FFC0&
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton CmdMbrEntry 
            BackColor       =   &H00FFFF80&
            Caption         =   "Member's Action Area"
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
            Left            =   480
            MaskColor       =   &H00C0FFC0&
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   600
            Width           =   2895
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pick Any From The Menu"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   2
         Top             =   960
         Width           =   9135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "What Do You Wanna Do ???"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   1
         Top             =   360
         Width           =   9135
      End
   End
   Begin VB.Menu madmin 
      Caption         =   "Admin"
      Index           =   1
   End
   Begin VB.Menu mmember 
      Caption         =   "Member"
      Index           =   2
   End
   Begin VB.Menu MnuBkIsuDtl 
      Caption         =   "Book"
      Index           =   3
   End
   Begin VB.Menu MnuUmg 
      Caption         =   "UserBase"
      Index           =   4
   End
   Begin VB.Menu MnuRpt 
      Caption         =   "Report"
      Index           =   5
   End
   Begin VB.Menu mabout 
      Caption         =   "About"
      Index           =   6
   End
End
Attribute VB_Name = "MDIfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdBkEntry_Click()
FrmBkEntry.Show
End Sub

Private Sub CmdBkRpt_Click()
FrmBookRpt.Show
End Sub

Private Sub CmdBkSubISu_Click()
FrmBookIsu.Show
End Sub

Private Sub CmdBookisland_Click()
FrmBookIsland.Show
End Sub

Private Sub CmdExit_Click()
Dim x As String
x = MsgBox("Do you really want to exit", vbYesNo + vbCritical, "Delete Confirmation")
If x = vbYes Then
End
End If

End Sub

Private Sub CmdIsuDtl_Click()
FrmBookIsuRu.Show

End Sub

Private Sub CmdIsuRpt_Click()
FrmIssueReport.Show

End Sub

Private Sub CmdMbrEntry_Click()
FrmMember.Show

End Sub

Private Sub CmdMbrRpt_Click()
FrmMemRep.Show

End Sub
Private Sub cmduserbase_Click()
FrmUserBase.Show
End Sub

Private Sub Command1_Click()
FrmBookSu.Show

End Sub

Private Sub Command7_Click()
FrmIssueReport.Show
End Sub

Private Sub Form_Resize()
    If Me.Width > 1000 And Me.Height > 1000 Then
        StatusBar1.Panels(1).Width = Me.ScaleWidth * 0.5
        StatusBar1.Panels(2).Width = Me.ScaleWidth * 0.11
        StatusBar1.Panels(3).Width = Me.ScaleWidth * 0.11
        StatusBar1.Panels(4).Width = Me.ScaleWidth * 0.11

StatusBar1.Panels(1) = "Current User : " & Book1.userNm & ""
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mabout_Click(Index As Integer)
FrmAbout.Show
End Sub

Private Sub madmin_Click(Index As Integer)
FrmAdminLogin.Show
End Sub

Private Sub mmember_Click(Index As Integer)
FrmMember.Show
End Sub

Private Sub MnuBkIsuDtl_Click(Index As Integer)
FrmBkEntry.Show
End Sub

Private Sub MnuRpt_Click(Index As Integer)
FrmMemRep.Show
End Sub

Private Sub MnuUmg_Click(Index As Integer)
FrmUserBase.Show
End Sub
