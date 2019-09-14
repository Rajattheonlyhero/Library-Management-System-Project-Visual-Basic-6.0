VERSION 5.00
Begin VB.Form FrmBookIsuRu 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Book issue Rules"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Library Rules"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   8655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   14175
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmBookIsuDe.frx":0000
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   8
         Top             =   7440
         Width           =   12855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "7.Working Hours of the Library: 24*7*365(Except for The Server Down for Maintenance)"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   7
         Top             =   6360
         Width           =   12855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "6.Students are allowed to library only on production of their authorized/valid Identity Cards"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   6
         Top             =   5400
         Width           =   12855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "5.Library borrower cards are not transferable. The borrower is responsible for the books borrowed on his/her card."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   5
         Top             =   4440
         Width           =   12855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "4.Before taking Any Book, Plzz Do fill the Issue Form.Any Illegal Means is Punishable."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   4
         Top             =   3480
         Width           =   12855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "3.Enter your UserName and Password and Sign In before entering library Portal."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   3
         Top             =   2520
         Width           =   12855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "2.Registration should be done to become a library member prior to using the library resources."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   2
         Top             =   1560
         Width           =   12855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1. User ID And Password is compulsory for getting access to the Library."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   12855
      End
   End
End
Attribute VB_Name = "FrmBookIsuRu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

